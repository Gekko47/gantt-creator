using ExcelDna.Integration.CustomUI;
using GanttCreator.Core;
using GanttCreator.Office;
using Moq;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// Contract tests for the R5.6a custom-style controls: the item list is the
/// configuration registry's user-created (non-built-in) keys in catalogue
/// order, the dropdown enables only for exactly one style-key-type body row
/// with at least one custom style existing, a built-in assignment reads as
/// no match, and the command re-checks the selection at execution (ruling 1).
/// </summary>
/// <remarks>
/// The class shares the <c>diagnostics-service</c> collection because
/// <see cref="RibbonStateService"/> is a session singleton armed by
/// <see cref="AddInHost.AutoOpen"/>; xUnit runs a collection's classes
/// sequentially, so the host lifecycle tests cannot race these.
/// </remarks>
[Collection("diagnostics-service")]
public sealed class CustomStyleRibbonTests : IDisposable
{
    public CustomStyleRibbonTests() => RibbonStateService.Reset();

    public void Dispose()
    {
        RibbonStateService.Reset();
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// Hand-written fake for <see cref="ISelectionIdentifier"/>: returns a
    /// fixed selection (the same shape the description-label tests use).
    /// </summary>
    private sealed class FakeSelectionIdentifier : ISelectionIdentifier
    {
        public SelectionFacts Facts { get; set; } = SelectionFacts.None;

        public SelectionFacts Identify() => Facts;

        public IDisposable SubscribeSelectionChanged(Action handler)
        {
            ArgumentNullException.ThrowIfNull(handler);
            return new Subscription();
        }

        private sealed class Subscription : IDisposable
        {
            public void Dispose()
            {
            }
        }
    }

    /// <summary>
    /// Minimal <see cref="IExcelApplicationAdapter"/> that reports a live
    /// workbook, so the shared gate's workbook half is satisfied and the
    /// tests isolate the custom-style-specific half.
    /// </summary>
    private sealed class FakeApplicationAdapter : IExcelApplicationAdapter
    {
        public bool? HasActiveWorkbook() => true;

        public void ShowStatusBar(string message)
        {
            // Intentionally empty: the status bar is not under test here.
        }

        public IDisposable SubscribeWorkbookStateChanged(Action handler) => new Subscription();

        private sealed class Subscription : IDisposable
        {
            public void Dispose()
            {
            }
        }
    }

    /// <summary>
    /// Builds one table body row through the normalised DTO shape; the style
    /// key sits in its normalised position so a test can pin a blank cell, a
    /// built-in assignment, or a custom assignment without touching Excel.
    /// </summary>
    private static GanttRowDto Row(string type, string? styleKey) => new(
        1,
        "G-" + new string('0', 31) + 'a',
        null,
        null,
        type,
        null,
        null,
        null,
        null,
        styleKey,
        null,
        null,
        null,
        null,
        null);

    /// <summary>
    /// The first code-owned built-in preset key, used to seed a registry
    /// entry that must NEVER appear in the dropdown.
    /// </summary>
    private static string BuiltInKey => GanttCatalogues.StylePresets[0].StyleKey;

    /// <summary>
    /// Arms the service with a single-selected body row of the given type and
    /// a registry carrying the given style keys (built-in key first, then the
    /// customs), and activates it.
    /// </summary>
    private static RibbonStateService ArmWithRow(
        FakeSelectionIdentifier selection,
        GanttRowDto row,
        params string[] customKeys)
    {
        var definitions = new List<GanttStyleDefinition>
        {
            new(
                BuiltInKey,
                new HashSet<GanttLabelPosition>([GanttLabelPosition.Auto]),
                EntityColourCapability.Fill),
        };
        foreach (string key in customKeys)
        {
            definitions.Add(new GanttStyleDefinition(
                key,
                new HashSet<GanttLabelPosition>([GanttLabelPosition.Auto]),
                EntityColourCapability.StyleDefined));
        }

        var service = RibbonStateService.Instance;
        service.SetApplicationAdapter(new FakeApplicationAdapter());
        service.SetSelectionIdentifier(selection);
        service.SetRibbon(null);

        var reader = new Mock<IGanttTableReader>();
        reader.Setup(item => item.Read()).Returns(GanttTableReadOutcome.Ok([row]));
        service.SetTableReader(reader.Object);

        var catalogues = new Mock<IConfigCatalogueReader>();
        catalogues.Setup(item => item.Read()).Returns(
            ConfigReadOutcome.Ok(
                "W-1",
                new Dictionary<string, string>(),
                new GanttStyleRegistry(definitions),
                new Dictionary<string, double>()));
        service.SetCatalogueReader(catalogues.Object);

        service.Activate();
        return service;
    }

    private static SelectionFacts SingleRow() => new(true, true, 1, false, 1);

    [Fact]
    public void The_dropdown_is_disabled_when_the_selection_is_not_one_body_row()
    {
        var selection = new FakeSelectionIdentifier
        {
            Facts = new SelectionFacts(true, false, null, true, 0),
        };
        var service = ArmWithRow(
            selection,
            Row(EntityTypeCatalog.Format(GanttEntityType.CustomActivity), null),
            "MyStyle");

        Assert.False(service.GetEnabled(RibbonControlIds.CustomStyle));
    }

    [Fact]
    public void The_dropdown_is_enabled_for_a_single_custom_activity_row_with_custom_styles()
    {
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var service = ArmWithRow(
            selection,
            Row(EntityTypeCatalog.Format(GanttEntityType.CustomActivity), null),
            "MyStyle");

        Assert.True(service.GetEnabled(RibbonControlIds.CustomStyle));
    }

    [Fact]
    public void The_dropdown_is_enabled_for_a_single_custom_milestone_row()
    {
        // R5.6a: Custom Milestone is a style-key type exactly as Custom
        // Activity is, so the assignment control enables for it too.
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var service = ArmWithRow(
            selection,
            Row(EntityTypeCatalog.Format(GanttEntityType.CustomMilestone), null),
            "MyStyle");

        Assert.True(service.GetEnabled(RibbonControlIds.CustomStyle));
    }

    [Fact]
    public void The_dropdown_is_disabled_when_no_custom_style_exists()
    {
        // Positive test for the empty-options gate: only built-ins are
        // registered, so the options are empty and the dropdown greys out
        // rather than presenting an empty list.
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var service = ArmWithRow(
            selection,
            Row(EntityTypeCatalog.Format(GanttEntityType.CustomActivity), null));

        Assert.Equal(1, service.GetCustomStyleCount()); // the leading <c>None</c> clear item
        Assert.False(service.GetEnabled(RibbonControlIds.CustomStyle));
    }

    [Fact]
    public void The_dropdown_is_disabled_for_a_non_style_type_even_with_custom_styles()
    {
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var service = ArmWithRow(
            selection,
            Row(EntityTypeCatalog.Format(GanttEntityType.AsPlannedActivity), null),
            "MyStyle");

        // The control is disabled because the selection is not on a
        // style-key row (Custom Activity / Custom Milestone), even though
        // custom styles exist in the catalogue. The dropdown's item list
        // always reflects the current catalogue: None + MyStyle = 2 items.
        Assert.False(service.GetEnabled(RibbonControlIds.CustomStyle));
        Assert.Equal(2, service.GetCustomStyleCount()); // None + MyStyle
    }

    [Fact]
    public void The_dropdown_offers_only_user_created_keys_in_catalogue_order()
    {
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var service = ArmWithRow(
            selection,
            Row(EntityTypeCatalog.Format(GanttEntityType.CustomActivity), null),
            "Alpha", "Beta");

        Assert.Equal(3, service.GetCustomStyleCount()); // <c>None</c> plus the two custom styles
        Assert.Equal("None", service.GetCustomStyleItemID(0));
        Assert.Equal("Alpha", service.GetCustomStyleItemID(1));
        Assert.Equal("Beta", service.GetCustomStyleItemID(2));
        Assert.Equal("None", service.GetCustomStyleItemLabel(0));
        Assert.Equal("Alpha", service.GetCustomStyleItemLabel(1));
        Assert.Equal("Beta", service.GetCustomStyleItemLabel(2));

        // The built-in key seeded into the registry is never offered.
        for (var index = 0; index < service.GetCustomStyleCount(); index++)
        {
            Assert.NotEqual(BuiltInKey, service.GetCustomStyleItemID(index));
        }
    }

    [Fact]
    public void The_leading_item_is_the_none_clear_item_and_count_includes_it()
    {
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };

        // Count always includes the leading <c>None</c> clear item.
        var withStyles = ArmWithRow(
            selection,
            Row(EntityTypeCatalog.Format(GanttEntityType.CustomActivity), null),
            "Alpha", "Beta");
        Assert.Equal(3, withStyles.GetCustomStyleCount());
        Assert.Equal("None", withStyles.GetCustomStyleItemID(0));
        Assert.Equal("None", withStyles.GetCustomStyleItemLabel(0));
        Assert.Equal("Alpha", withStyles.GetCustomStyleItemID(1));
        Assert.Equal("Beta", withStyles.GetCustomStyleItemID(2));

        // No custom styles: only <c>None</c> is offered, so the dropdown never
        // presents an empty list that Excel would misread as index 0.
        var noneOnly = ArmWithRow(
            selection,
            Row(EntityTypeCatalog.Format(GanttEntityType.CustomActivity), null));
        Assert.Equal(1, noneOnly.GetCustomStyleCount());
        Assert.Equal("None", noneOnly.GetCustomStyleItemID(0));
        Assert.Equal("None", noneOnly.GetCustomStyleItemLabel(0));

        // A blank style key selects <c>None</c> (index 0).
        var blank = ArmWithRow(
            selection,
            Row(EntityTypeCatalog.Format(GanttEntityType.CustomActivity), null),
            "Alpha");
        Assert.Equal(0, blank.GetCustomStyleSelectedIndex());
    }

    [Fact]
    public void Out_of_range_indices_yield_empty_items()
    {
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var service = ArmWithRow(
            selection,
            Row(EntityTypeCatalog.Format(GanttEntityType.CustomActivity), null),
            "MyStyle");

        Assert.Equal(string.Empty, service.GetCustomStyleItemID(-1));
        Assert.Equal(string.Empty, service.GetCustomStyleItemID(2)); // out of range: <c>None</c> + one custom style
        Assert.Equal(string.Empty, service.GetCustomStyleItemLabel(5));
    }

    [Fact]
    public void The_selected_index_matches_the_row_custom_key()
    {
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var service = ArmWithRow(
            selection,
            Row(EntityTypeCatalog.Format(GanttEntityType.CustomActivity), "Beta"),
            "Alpha", "Beta");

        Assert.Equal(2, service.GetCustomStyleSelectedIndex()); // "Beta" is options[1], so it projects to index 2 (<c>None</c> leads)
    }

    [Fact]
    public void The_selected_index_is_zero_for_a_blank_or_builtin_style_key()
    {
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var blank = ArmWithRow(
            selection,
            Row(EntityTypeCatalog.Format(GanttEntityType.CustomActivity), null),
            "MyStyle");
        Assert.Equal(0, blank.GetCustomStyleSelectedIndex());

        RibbonStateService.Reset();

        var builtIn = ArmWithRow(
            selection,
            Row(EntityTypeCatalog.Format(GanttEntityType.CustomActivity), BuiltInKey),
            "MyStyle");
        Assert.Equal(0, builtIn.GetCustomStyleSelectedIndex());
    }

    [Fact]
    public void A_style_created_after_the_capture_appears_once_the_ribbon_state_is_refreshed()
    {
        // R5.6a: the dropdown's items are projected from the workbook's style
        // registry at capture time. A style the user creates in the Theme
        // settings dialog is only reachable once the ribbon state has been
        // re-captured, which is what the post-command hook does.
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var service = ArmWithRow(
            selection,
            Row(EntityTypeCatalog.Format(GanttEntityType.CustomActivity), null));

        Assert.Equal(1, service.GetCustomStyleCount()); // only the leading <c>None</c> item

        // The user creates a style in the Theme settings dialog: the workbook
        // now carries it, so the next read projects it.
        var created = new List<GanttStyleDefinition>
        {
            new(
                BuiltInKey,
                new HashSet<GanttLabelPosition>([GanttLabelPosition.Auto]),
                EntityColourCapability.Fill),
            new(
                "Created",
                new HashSet<GanttLabelPosition>([GanttLabelPosition.Auto]),
                EntityColourCapability.StyleDefined),
        };
        var catalogues = new Mock<IConfigCatalogueReader>();
        catalogues.Setup(item => item.Read()).Returns(
            ConfigReadOutcome.Ok(
                "W-1",
                new Dictionary<string, string>(),
                new GanttStyleRegistry(created),
                new Dictionary<string, double>()));
        service.SetCatalogueReader(catalogues.Object);

        service.Refresh();

        Assert.Equal(2, service.GetCustomStyleCount());
        Assert.Equal("Created", service.GetCustomStyleItemID(1));
    }

    [Fact]
    public void The_ribbon_callbacks_route_through_the_state_service()
    {
        // Excel dispatches these by reflection, so a rename in either layer is
        // a live defect; the routing overloads pin the wiring without the
        // session singleton.
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var service = ArmWithRow(
            selection,
            Row(EntityTypeCatalog.Format(GanttEntityType.CustomActivity), null),
            "MyStyle");
        var control = new Mock<IRibbonControl>().Object;

        Assert.Equal(service.GetCustomStyleCount(), GanttRibbon.GetCustomStyleCount(control, service));
        Assert.Equal(service.GetCustomStyleItemID(0), GanttRibbon.GetCustomStyleItemID(control, 0, service));
        Assert.Equal(service.GetCustomStyleItemLabel(0), GanttRibbon.GetCustomStyleItemLabel(control, 0, service));
        Assert.Equal(
            service.GetCustomStyleSelectedIndex(),
            GanttRibbon.GetCustomStyleSelectedIndex(control, service));
    }

    [Fact]
    public void A_null_state_service_is_refused_by_every_custom_style_callback()
    {
        var control = new Mock<IRibbonControl>().Object;

        Assert.Throws<ArgumentNullException>(() => GanttRibbon.GetCustomStyleCount(control, null!));
        Assert.Throws<ArgumentNullException>(() => GanttRibbon.GetCustomStyleItemID(control, 0, null!));
        Assert.Throws<ArgumentNullException>(() => GanttRibbon.GetCustomStyleItemLabel(control, 0, null!));
        Assert.Throws<ArgumentNullException>(() => GanttRibbon.GetCustomStyleSelectedIndex(control, null!));
        Assert.Throws<ArgumentNullException>(() => GanttRibbon.OnCustomStyleChange(control, "MyStyle", null!));
    }

    [Fact]
    public void An_unarmed_service_reports_an_empty_dropdown_rather_than_throwing()
    {
        var control = new Mock<IRibbonControl>().Object;
        var service = new RibbonStateService();

        Assert.Equal(1, GanttRibbon.GetCustomStyleCount(control, service)); // the leading <c>None</c> item
        Assert.Equal(0, GanttRibbon.GetCustomStyleSelectedIndex(control, service));
        Assert.Equal("None", GanttRibbon.GetCustomStyleItemLabel(control, 0, service));
        Assert.Equal("None", GanttRibbon.GetCustomStyleItemID(control, 0, service));
    }

    [Fact]
    public void The_command_refuses_a_non_single_row_selection_before_the_mutator_runs()
    {
        // Ruling 1 at execution time: the ribbon gate may have been stale, so
        // the command re-reads the selection and refuses before the write.
        var mutator = new Mock<IStyleKeyMutator>();
        var selection = new Mock<ISelectionIdentifier>();
        selection.Setup(item => item.Identify())
            .Returns(new SelectionFacts(true, false, null, true, 0));
        var presented = new List<string>();

        SetStyleKeyCommand.Run(mutator.Object, selection.Object, presented.Add, "MyStyle");

        mutator.Verify(item => item.SetStyleKey(It.IsAny<string>()), Times.Never);
        Assert.Single(presented);
        Assert.Contains("data row", presented[0], StringComparison.Ordinal);
    }

    [Fact]
    public void The_command_silence_on_success()
    {
        var mutator = new Mock<IStyleKeyMutator>();
        mutator.Setup(item => item.SetStyleKey("MyStyle"))
            .Returns(StyleKeyMutationOutcome.Ok(1, 1));
        var selection = new Mock<ISelectionIdentifier>();
        selection.Setup(item => item.Identify()).Returns(SingleRow());
        var presented = new List<string>();

        SetStyleKeyCommand.Run(mutator.Object, selection.Object, presented.Add, "MyStyle");

        mutator.Verify(item => item.SetStyleKey("MyStyle"), Times.Once);
        Assert.Empty(presented);
    }

    [Fact]
    public void The_command_presents_the_translated_refusal()
    {
        var mutator = new Mock<IStyleKeyMutator>();
        mutator.Setup(item => item.SetStyleKey(It.IsAny<string>()))
            .Returns(StyleKeyMutationOutcome.Refused(
                StyleKeyMutationRefusalReason.TypeNotCustomStyled, 2));
        var selection = new Mock<ISelectionIdentifier>();
        selection.Setup(item => item.Identify()).Returns(SingleRow());
        var presented = new List<string>();

        SetStyleKeyCommand.Run(mutator.Object, selection.Object, presented.Add, "MyStyle");

        var message = Assert.Single(presented);
        Assert.StartsWith("Row 2 ", message, StringComparison.Ordinal);
        Assert.Contains("Custom Activity and Custom Milestone", message, StringComparison.Ordinal);
    }

    [Fact]
    public void A_throwing_presenter_does_not_escape_the_command()
    {
        var mutator = new Mock<IStyleKeyMutator>();
        mutator.Setup(item => item.SetStyleKey(It.IsAny<string>()))
            .Returns(StyleKeyMutationOutcome.Refused(StyleKeyMutationRefusalReason.WriteFailed, 1));
        var selection = new Mock<ISelectionIdentifier>();
        selection.Setup(item => item.Identify()).Returns(SingleRow());

        var exception = Record.Exception(() => SetStyleKeyCommand.Run(
            mutator.Object,
            selection.Object,
            _ => throw new InvalidOperationException("dialog refused"),
            "MyStyle"));

        Assert.Null(exception);
    }
}
