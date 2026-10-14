using ExcelDna.Integration.CustomUI;
using GanttCreator.Core;
using GanttCreator.Office;
using Moq;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// Contract tests for the R5.5 label-position controls (work item
/// <c>R5.5-label-controls.md</c>): the item list is projected from the selected
/// row's <c>Type</c> (D1), the wording table never shows an enum name verbatim
/// (ruling 3), the dropdown enables only for exactly one body row (ruling 6), a
/// blank cell reads as <c>Default</c> (ruling 5), and an unknown item ID is
/// refused rather than defaulted.
/// </summary>
/// <remarks>
/// The class shares the <c>diagnostics-service</c> collection because
/// <see cref="RibbonStateService"/> is a session singleton armed by
/// <see cref="AddInHost.AutoOpen"/>; xUnit runs a collection's classes
/// sequentially, so the host lifecycle tests cannot race these.
/// </remarks>
[Collection("diagnostics-service")]
public sealed class LabelPositionRibbonTests : IDisposable
{
    public LabelPositionRibbonTests() => RibbonStateService.Reset();

    public void Dispose()
    {
        RibbonStateService.Reset();
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// Hand-written fake for <see cref="ISelectionIdentifier"/>: returns a fixed
    /// selection and counts subscription attach/detach so the tests can prove
    /// only the label dropdown is invalidated on a selection change.
    /// </summary>
    private sealed class FakeSelectionIdentifier : ISelectionIdentifier
    {
        private Action? _handler;

        public SelectionFacts Facts { get; set; } = SelectionFacts.None;

        public int SubscribeCount { get; private set; }

        public int DisposeCount { get; private set; }

        public SelectionFacts Identify() => Facts;

        public IDisposable SubscribeSelectionChanged(Action handler)
        {
            ArgumentNullException.ThrowIfNull(handler);
            SubscribeCount++;
            _handler = handler;
            return new Subscription(this);
        }

        internal void RaiseSelectionChanged() => _handler?.Invoke();

        private sealed class Subscription(FakeSelectionIdentifier owner) : IDisposable
        {
            private bool _disposed;

            public void Dispose()
            {
                if (_disposed)
                {
                    return;
                }

                _disposed = true;
                owner.DisposeCount++;
                owner._handler = null;
            }
        }
    }

    /// <summary>
    /// Builds one table body row through the normalised DTO shape, keeping the
    /// raw cell text the projection reads, so a test can pin a blank
    /// <c>LabelPosition</c> or an unknown <c>Type</c> without touching Excel.
    /// </summary>
    private static GanttRowDto Row(string type, string? labelPosition) => new(
        1,
        "G-" + new string('0', 31) + 'a',
        null,
        null,
        type,
        null,
        null,
        null,
        null,
        null,
        labelPosition,
        null,
        null,
        null,
        null);

    /// <summary>
    /// Minimal <see cref="IExcelApplicationAdapter"/> that reports a live
    /// workbook, so the shared gate's workbook half is satisfied and the tests
    /// isolate the label-specific half (single body row).
    /// </summary>
    private sealed class FakeApplicationAdapter : IExcelApplicationAdapter
    {
        public bool? HasActiveWorkbook() => true;

        public IDisposable SubscribeWorkbookStateChanged(Action handler) => new Subscription();

        private sealed class Subscription : IDisposable
        {
            public void Dispose()
            {
            }
        }
    }

    /// <summary>
    /// Arms the service with a single-selected body row of the given type and
    /// activates it, so the label getters read a resolved context and the
    /// selection subscription is attached exactly as AutoOpen attaches it.
    /// </summary>
    private static RibbonStateService ArmWithRow(
        FakeSelectionIdentifier selection,
        GanttRowDto row,
        GanttStyleRegistry? styles = null,
        IRibbonUI? ribbon = null)
    {
        var service = RibbonStateService.Instance;
        service.SetApplicationAdapter(new FakeApplicationAdapter());
        service.SetSelectionIdentifier(selection);
        service.SetRibbon(ribbon);

        var reader = new Mock<IGanttTableReader>();
        reader.Setup(item => item.Read()).Returns(GanttTableReadOutcome.Ok([row]));
        service.SetTableReader(reader.Object);

        var catalogues = new Mock<IConfigCatalogueReader>();
        catalogues.Setup(item => item.Read()).Returns(
            ConfigReadOutcome.Ok(
                "W-1",
                new Dictionary<string, string>(),
                styles ?? GanttStyleRegistry.Empty));
        service.SetCatalogueReader(catalogues.Object);

        service.Activate();
        return service;
    }

    private static SelectionFacts SingleRow() => new(true, true, 1, false, 1);

    [Fact]
    public void The_dropdown_is_disabled_when_the_selection_is_not_one_body_row()
    {
        // Ruling 1 + 6: a header, multi, or absent selection disables the
        // control, so a group of mixed types can never be formatted at once.
        var selection = new FakeSelectionIdentifier
        {
            Facts = new SelectionFacts(true, false, null, true, 0),
        };
        var service = ArmWithRow(selection, Row("As-Built Activity", null));

        Assert.False(service.GetEnabled(RibbonControlIds.LabelPosition));
    }

    [Fact]
    public void The_dropdown_is_enabled_for_exactly_one_body_row()
    {
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var service = ArmWithRow(selection, Row("As-Built Activity", null));

        Assert.True(service.GetEnabled(RibbonControlIds.LabelPosition));
    }

    [Fact]
    public void The_dropdown_offers_the_span_set_plus_default_for_a_span_type()
    {
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var service = ArmWithRow(selection, Row("As-Built Activity", null));

        // Default (clear) is always offered first, then the catalogue set in
        // display order — D1 projection, never a hand-authored list.
        Assert.Equal(6, service.GetLabelPositionCount());
        Assert.Equal("Default", service.GetLabelPositionItemID(0));
        Assert.Equal("Auto", service.GetLabelPositionItemID(1));
        Assert.Equal("Left", service.GetLabelPositionItemID(2));
        Assert.Equal("Right", service.GetLabelPositionItemID(3));
        Assert.Equal("Inside", service.GetLabelPositionItemID(4));
        Assert.Equal("None", service.GetLabelPositionItemID(5));
    }

    [Fact]
    public void The_dropdown_offers_no_auto_for_a_splitter()
    {
        // The splitter has no Auto in its catalogue set, and the projection
        // must not invent one.
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var service = ArmWithRow(selection, Row("Splitter", null));

        Assert.Equal(5, service.GetLabelPositionCount());
        Assert.Equal("DataPanelLeft", service.GetLabelPositionItemID(1));
        Assert.Equal("PlotCentre", service.GetLabelPositionItemID(2));
        Assert.Equal("Both", service.GetLabelPositionItemID(3));
        Assert.Equal("None", service.GetLabelPositionItemID(4));
    }

    [Theory]
    [InlineData("Spacer")]
    [InlineData("Critical Interval")]
    public void Single_value_types_offer_only_default_and_none_hide(string type)
    {
        // Ruling 4: the catalogue set stays None-only. Left/Right are refused
        // by the projection and by the mutator, not silently offered.
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var service = ArmWithRow(selection, Row(type, null));

        Assert.Equal(2, service.GetLabelPositionCount());
        Assert.Equal("Default", service.GetLabelPositionItemID(0));
        Assert.Equal("None", service.GetLabelPositionItemID(1));
    }

    [Fact]
    public void The_dropdown_shows_the_ruling_3_wording_not_the_enum_name()
    {
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var service = ArmWithRow(selection, Row("As-Built Activity", null));

        // Inside reads Centre (R4.7G D6) and None reads None (hide), so the
        // explicit hide stays distinct from Default.
        Assert.Equal("Default", service.GetLabelPositionItemLabel(0));
        Assert.Equal("Centre", service.GetLabelPositionItemLabel(4));
        Assert.Equal("None (hide)", service.GetLabelPositionItemLabel(5));
    }

    [Fact]
    public void The_dropdown_labels_the_splitter_positions()
    {
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var service = ArmWithRow(selection, Row("Splitter", null));

        Assert.Equal("Data panel", service.GetLabelPositionItemLabel(1));
        Assert.Equal("Plot centre", service.GetLabelPositionItemLabel(2));
        Assert.Equal("Both", service.GetLabelPositionItemLabel(3));
    }

    [Fact]
    public void The_dropdown_labels_the_delineator_corners()
    {
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var service = ArmWithRow(selection, Row("Delineator", null));

        Assert.Equal(7, service.GetLabelPositionCount());
        Assert.Equal("Top left", service.GetLabelPositionItemLabel(2));
        Assert.Equal("Bottom right", service.GetLabelPositionItemLabel(5));
        Assert.Equal("None (hide)", service.GetLabelPositionItemLabel(6));
    }

    [Fact]
    public void A_blank_cell_reads_as_default()
    {
        // Ruling 5: blank resolves to the style default / Auto cascade (§22),
        // so the dropdown selects the leading Default item.
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var service = ArmWithRow(selection, Row("As-Built Activity", null));

        Assert.Equal(0, service.GetLabelPositionSelectedIndex());
    }

    [Fact]
    public void A_stored_position_selects_its_own_item()
    {
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var service = ArmWithRow(selection, Row("As-Built Activity", "Right"));

        Assert.Equal(3, service.GetLabelPositionSelectedIndex());
    }

    [Fact]
    public void An_out_of_range_item_index_returns_an_empty_label()
    {
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var service = ArmWithRow(selection, Row("As-Built Activity", null));

        Assert.Equal(string.Empty, service.GetLabelPositionItemLabel(99));
        Assert.Equal(string.Empty, service.GetLabelPositionItemID(99));
    }

    [Fact]
    public void No_resolved_selection_offers_no_items_at_all()
    {
        // An empty list is Excel's cue that the dropdown has nothing to show,
        // and it pairs with the disabled gate rather than an empty picker.
        var selection = new FakeSelectionIdentifier { Facts = SelectionFacts.None };
        var service = ArmWithRow(selection, Row("As-Built Activity", null));

        Assert.Equal(0, service.GetLabelPositionCount());
        Assert.Equal(0, service.GetLabelPositionSelectedIndex());
    }

    [Fact]
    public void An_unknown_type_leaves_the_context_unavailable()
    {
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var service = ArmWithRow(selection, Row("Widget", null));

        Assert.Equal(0, service.GetLabelPositionCount());
        Assert.False(service.GetEnabled(RibbonControlIds.LabelPosition));
    }

    [Fact]
    public void A_collapsed_parent_row_still_resolves_its_own_label_context()
    {
        // Collapse clarification: a parent owning hidden children is still
        // exactly one body row, so the control enables and only that row's
        // cell is targeted.
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var service = ArmWithRow(selection, Row("As-Built Activity", null));

        Assert.True(service.GetEnabled(RibbonControlIds.LabelPosition));
        Assert.Equal(6, service.GetLabelPositionCount());
    }

    [Fact]
    public void A_selection_change_invalidates_only_the_label_dropdown()
    {
        // Ruling 6: moving the selection must not re-query every getter on the
        // ribbon, so the handler invalidates exactly the one control.
        var ribbon = new Mock<IRibbonUI>();
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        ArmWithRow(selection, Row("As-Built Activity", null), ribbon: ribbon.Object);

        Assert.Equal(1, selection.SubscribeCount);

        // Activation already invalidated the whole ribbon once; the
        // selection handler must not add to that.
        var invalidationsBefore = CountInvalidations(ribbon);

        selection.RaiseSelectionChanged();

        var invalidated = ribbon.Invocations
            .Where(item => item.Method.Name == nameof(IRibbonUI.InvalidateControl))
            .Select(item => item.Arguments[0])
            .ToList();

        Assert.Equal([RibbonControlIds.LabelPosition], invalidated);
        Assert.Equal(invalidationsBefore, CountInvalidations(ribbon));
    }

    [Fact]
    public void The_selection_subscription_is_detached_exactly_once_on_reset()
    {
        // R1.6: unload must leave no COM connection behind.
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        ArmWithRow(selection, Row("As-Built Activity", null));

        RibbonStateService.Reset();

        Assert.Equal(1, selection.DisposeCount);
    }

    [Fact]
    public void The_ribbon_callbacks_route_through_the_state_service()
    {
        // Excel dispatches these by reflection, so a rename in either layer is
        // a live defect; the routing overloads pin the wiring without the
        // session singleton.
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var service = ArmWithRow(selection, Row("As-Built Activity", null));
        var control = new Mock<IRibbonControl>().Object;

        Assert.Equal(service.GetLabelPositionCount(), GanttRibbon.GetLabelPositionCount(control, service));
        Assert.Equal(service.GetLabelPositionItemID(4), GanttRibbon.GetLabelPositionItemID(control, 4, service));
        Assert.Equal(service.GetLabelPositionItemLabel(4), GanttRibbon.GetLabelPositionItemLabel(control, 4, service));
        Assert.Equal(
            service.GetLabelPositionSelectedIndex(),
            GanttRibbon.GetLabelPositionSelectedIndex(control, service));
    }

    [Fact]
    public void A_null_state_service_is_refused_by_every_label_callback()
    {
        // The routing overloads validate their collaborator, so a wiring
        // mistake surfaces as a clear exception rather than a silent empty
        // dropdown in Excel.
        var control = new Mock<IRibbonControl>().Object;

        Assert.Throws<ArgumentNullException>(() => GanttRibbon.GetLabelPositionCount(control, null!));
        Assert.Throws<ArgumentNullException>(() => GanttRibbon.GetLabelPositionItemID(control, 0, null!));
        Assert.Throws<ArgumentNullException>(() => GanttRibbon.GetLabelPositionItemLabel(control, 0, null!));
        Assert.Throws<ArgumentNullException>(() => GanttRibbon.GetLabelPositionSelectedIndex(control, null!));
        Assert.Throws<ArgumentNullException>(() => GanttRibbon.OnLabelPositionChange(control, "Left", null!));
    }

    [Fact]
    public void An_unarmed_service_reports_an_empty_dropdown_rather_than_throwing()
    {
        // With no selection identifier or table reader the context cannot be
        // resolved, so the count is zero and no item index resolves: Excel
        // never queries an item because it was told there are none.
        var control = new Mock<IRibbonControl>().Object;
        var service = new RibbonStateService();

        Assert.Equal(0, GanttRibbon.GetLabelPositionCount(control, service));
        Assert.Equal(0, GanttRibbon.GetLabelPositionSelectedIndex(control, service));
        Assert.Equal(string.Empty, GanttRibbon.GetLabelPositionItemLabel(control, 1, service));
        Assert.Equal(string.Empty, GanttRibbon.GetLabelPositionItemID(control, 1, service));
    }

    private static int CountInvalidations(Mock<IRibbonUI> ribbon) =>
        ribbon.Invocations.Count(item => item.Method.Name == nameof(IRibbonUI.Invalidate));
}
