using ExcelDna.Integration.CustomUI;
using GanttCreator.Core;
using GanttCreator.Office;
using Moq;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// Contract tests for the R5.14 date-label controls (work item
/// <c>R5.14-date-labels.md</c>): the item list is projected from the
/// selected row's <c>Type</c> date mode (D1/D4) and already includes
/// <c>Auto</c>, so the leading Auto item carries the <c>Default</c>
/// clear-ID rather than a separate leading clear item; the wording
/// table never shows an enum name verbatim, the dropdown enables only
/// for exactly one body row of a dated type, a blank cell reads as
/// <c>Auto</c>, and an unknown item ID is refused rather than
/// defaulted.
/// </summary>
/// <remarks>
/// The class shares the <c>diagnostics-service</c> collection because
/// <see cref="RibbonStateService"/> is a session singleton armed by
/// <see cref="AddInHost.AutoOpen"/>; xUnit runs a collection's classes
/// sequentially, so the host lifecycle tests cannot race these.
/// </remarks>
[Collection("diagnostics-service")]
public sealed class DateLabelRibbonTests : IDisposable
{
    public DateLabelRibbonTests() => RibbonStateService.Reset();

    public void Dispose()
    {
        RibbonStateService.Reset();
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// Hand-written fake for <see cref="ISelectionIdentifier"/>: returns a fixed
    /// selection and counts subscription attach/detach so the tests can prove
    /// only the date-label dropdown is invalidated on a selection change.
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
    /// <c>DateLabelPosition</c> or an unknown <c>Type</c> without touching Excel.
    /// </summary>
    private static GanttRowDto Row(string type, string? dateLabelPosition) => new(
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
        null,
        null,
        null,
        null,
        null,
        dateLabelPosition);

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
    /// activates it, so the date-label getters read a resolved context and the
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
                styles ?? GanttStyleRegistry.Empty, new Dictionary<string, double>()));;
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

        Assert.False(service.GetEnabled(RibbonControlIds.DateLabel));
    }

    [Fact]
    public void The_dropdown_is_enabled_for_exactly_one_body_row()
    {
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var service = ArmWithRow(selection, Row("As-Built Activity", null));

        Assert.True(service.GetEnabled(RibbonControlIds.DateLabel));
    }

    [Fact]
    public void The_dropdown_offers_the_full_set_for_a_span_type()
    {
        // D1/D4: a span type names every choice in date order. The projected
        // set already includes Auto, so the leading item carries the Default
        // clear-ID (selecting it clears the cell to blank = Auto) and there is
        // no separate leading clear item as the description-label dropdown has.
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var service = ArmWithRow(selection, Row("As-Built Activity", null));

        Assert.Equal(4, service.GetDateLabelCount());
        Assert.Equal("Default", service.GetDateLabelItemID(0));
        Assert.Equal("Start", service.GetDateLabelItemID(1));
        Assert.Equal("Finish", service.GetDateLabelItemID(2));
        Assert.Equal("None", service.GetDateLabelItemID(3));
        Assert.Equal("Auto", service.GetDateLabelItemLabel(0));
        Assert.Equal("Start", service.GetDateLabelItemLabel(1));
        Assert.Equal("Finish", service.GetDateLabelItemLabel(2));
        Assert.Equal("None", service.GetDateLabelItemLabel(3));
    }

    [Fact]
    public void The_dropdown_omits_finish_for_a_start_only_type()
    {
        // A StartOnly type (milestone) has no Finish date, so
        // Finish is not offered: only Auto, Start, and None, in
        // that order. The delineator is StartOnly too, but its
        // date label is positionable (owner ruling 2026-10-14),
        // so it projects the corner set instead — see
        // The_dropdown_offers_the_six_corners_for_a_delineator_row.
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var service = ArmWithRow(selection, Row("As-Built Milestone", null));

        Assert.Equal(3, service.GetDateLabelCount());
        Assert.Equal("Default", service.GetDateLabelItemID(0));
        Assert.Equal("Start", service.GetDateLabelItemID(1));
        Assert.Equal("None", service.GetDateLabelItemID(2));
        Assert.Equal("Auto", service.GetDateLabelItemLabel(0));
        Assert.Equal("Start", service.GetDateLabelItemLabel(1));
        Assert.Equal("None", service.GetDateLabelItemLabel(2));
    }

    [Fact]
    public void A_dateless_type_offers_no_items_and_disables_the_dropdown()
    {
        // A dateless type (Splitter, Spacer) names no date-label choice at
        // all: the projection is empty, the context is unavailable, and the
        // dropdown is disabled rather than showing an empty picker.
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var service = ArmWithRow(selection, Row("Splitter", null));

        Assert.Equal(0, service.GetDateLabelCount());
        Assert.False(service.GetEnabled(RibbonControlIds.DateLabel));
    }

    [Fact]
    public void A_blank_cell_reads_as_auto()
    {
        // R5.14 D3: blank resolves to Auto (both date labels the type uses),
        // so the dropdown selects the leading Auto item.
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var service = ArmWithRow(selection, Row("As-Built Activity", null));

        Assert.Equal(0, service.GetDateLabelSelectedIndex());
    }

    [Fact]
    public void A_stored_start_selects_its_own_item()
    {
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var service = ArmWithRow(selection, Row("As-Built Activity", "Start"));

        Assert.Equal(1, service.GetDateLabelSelectedIndex());
    }

    [Fact]
    public void A_stored_finish_selects_its_own_item()
    {
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var service = ArmWithRow(selection, Row("As-Built Activity", "Finish"));

        Assert.Equal(2, service.GetDateLabelSelectedIndex());
    }

    [Fact]
    public void A_stored_none_selects_its_own_item()
    {
        // D5 (second owner ruling 2026-10-14): a stored None
        // (no date labels) selects its own item, past Finish.
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var service = ArmWithRow(selection, Row("As-Built Activity", "None"));

        Assert.Equal(3, service.GetDateLabelSelectedIndex());
    }

    [Fact]
    public void The_dropdown_offers_the_six_corners_for_a_delineator_row()
    {
        // Owner ruling 2026-10-14: a delineator's date label
        // is positionable, so the dropdown projects the
        // description label's corner set — not the which-date
        // choices a span offers. Auto leads and carries the
        // Default clear-ID, exactly as in the which-date
        // vocabulary; the wording is the shared label-position
        // authority, never the enum name.
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var service = ArmWithRow(selection, Row(EntityTypeCatalog.Format(GanttEntityType.Delineator), null));

        Assert.True(service.GetEnabled(RibbonControlIds.DateLabel));
        Assert.Equal(6, service.GetDateLabelCount());
        Assert.Equal("Default", service.GetDateLabelItemID(0));
        Assert.Equal("TopLeft", service.GetDateLabelItemID(1));
        Assert.Equal("TopRight", service.GetDateLabelItemID(2));
        Assert.Equal("BottomLeft", service.GetDateLabelItemID(3));
        Assert.Equal("BottomRight", service.GetDateLabelItemID(4));
        Assert.Equal("None", service.GetDateLabelItemID(5));
        Assert.Equal("Auto", service.GetDateLabelItemLabel(0));
        Assert.Equal("Top left", service.GetDateLabelItemLabel(1));
        Assert.Equal("Top right", service.GetDateLabelItemLabel(2));
        Assert.Equal("Bottom left", service.GetDateLabelItemLabel(3));
        Assert.Equal("Bottom right", service.GetDateLabelItemLabel(4));
        Assert.Equal("None (hide)", service.GetDateLabelItemLabel(5));
    }

    [Fact]
    public void A_blank_delineator_cell_selects_the_leading_auto_item()
    {
        // The placement vocabulary's leading choice is Auto,
        // exactly as the which-date vocabulary's is, so a
        // blank cell reads as index 0 in either.
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var service = ArmWithRow(selection, Row(EntityTypeCatalog.Format(GanttEntityType.Delineator), null));

        Assert.Equal(0, service.GetDateLabelSelectedIndex());
    }

    [Fact]
    public void A_stored_corner_selects_its_own_item()
    {
        // A stored corner name selects its own place in the
        // projected set — TopLeft is the first corner past
        // Auto, BottomRight the last before None.
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var topLeft = ArmWithRow(selection, Row(EntityTypeCatalog.Format(GanttEntityType.Delineator), "TopLeft"));
        Assert.Equal(1, topLeft.GetDateLabelSelectedIndex());

        var bottomRight = ArmWithRow(selection, Row(EntityTypeCatalog.Format(GanttEntityType.Delineator), "BottomRight"));
        Assert.Equal(4, bottomRight.GetDateLabelSelectedIndex());
    }

    [Fact]
    public void An_out_of_range_item_index_returns_an_empty_label()
    {
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var service = ArmWithRow(selection, Row("As-Built Activity", null));

        Assert.Equal(string.Empty, service.GetDateLabelItemLabel(99));
        Assert.Equal(string.Empty, service.GetDateLabelItemID(99));
    }

    [Fact]
    public void No_resolved_selection_offers_no_items_at_all()
    {
        // An empty list is Excel's cue that the dropdown has nothing to show,
        // and it pairs with the disabled gate rather than an empty picker.
        var selection = new FakeSelectionIdentifier { Facts = SelectionFacts.None };
        var service = ArmWithRow(selection, Row("As-Built Activity", null));

        Assert.Equal(0, service.GetDateLabelCount());
        Assert.Equal(0, service.GetDateLabelSelectedIndex());
    }

    [Fact]
    public void An_unknown_type_leaves_the_context_unavailable()
    {
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var service = ArmWithRow(selection, Row("Widget", null));

        Assert.Equal(0, service.GetDateLabelCount());
        Assert.False(service.GetEnabled(RibbonControlIds.DateLabel));
    }

    [Fact]
    public void A_collapsed_parent_row_still_resolves_its_own_date_label_context()
    {
        // Collapse clarification: a parent owning hidden children is still
        // exactly one body row, so the control enables and only that row's
        // cell is targeted.
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var service = ArmWithRow(selection, Row("As-Built Activity", null));

        Assert.True(service.GetEnabled(RibbonControlIds.DateLabel));
        Assert.Equal(4, service.GetDateLabelCount());
    }

    [Fact]
    public void A_selection_change_invalidates_only_the_two_label_dropdowns()
    {
        // Ruling 6: moving the selection must not re-query every getter on
        // the ribbon, so the handler invalidates exactly the two label
        // dropdowns that re-project from the newly selected row — not the
        // whole ribbon, and not any other control.
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

        Assert.Equal([RibbonControlIds.DescriptionLabel, RibbonControlIds.DateLabel], invalidated);
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

        Assert.Equal(service.GetDateLabelCount(), GanttRibbon.GetDateLabelCount(control, service));
        Assert.Equal(service.GetDateLabelItemID(2), GanttRibbon.GetDateLabelItemID(control, 2, service));
        Assert.Equal(service.GetDateLabelItemLabel(2), GanttRibbon.GetDateLabelItemLabel(control, 2, service));
        Assert.Equal(
            service.GetDateLabelSelectedIndex(),
            GanttRibbon.GetDateLabelSelectedIndex(control, service));
    }

    [Fact]
    public void A_null_state_service_is_refused_by_every_date_label_callback()
    {
        // The routing overloads validate their collaborator, so a wiring
        // mistake surfaces as a clear exception rather than a silent empty
        // dropdown in Excel.
        var control = new Mock<IRibbonControl>().Object;

        Assert.Throws<ArgumentNullException>(() => GanttRibbon.GetDateLabelCount(control, null!));
        Assert.Throws<ArgumentNullException>(() => GanttRibbon.GetDateLabelItemID(control, 0, null!));
        Assert.Throws<ArgumentNullException>(() => GanttRibbon.GetDateLabelItemLabel(control, 0, null!));
        Assert.Throws<ArgumentNullException>(() => GanttRibbon.GetDateLabelSelectedIndex(control, null!));
        Assert.Throws<ArgumentNullException>(() => GanttRibbon.OnDateLabelChange(control, "Start", null!));
    }

    [Fact]
    public void An_unarmed_service_reports_an_empty_dropdown_rather_than_throwing()
    {
        // With no selection identifier or table reader the context cannot be
        // resolved, so the count is zero and no item index resolves: Excel
        // never queries an item because it was told there are none.
        var control = new Mock<IRibbonControl>().Object;
        var service = new RibbonStateService();

        Assert.Equal(0, GanttRibbon.GetDateLabelCount(control, service));
        Assert.Equal(0, GanttRibbon.GetDateLabelSelectedIndex(control, service));
        Assert.Equal(string.Empty, GanttRibbon.GetDateLabelItemLabel(control, 1, service));
        Assert.Equal(string.Empty, GanttRibbon.GetDateLabelItemID(control, 1, service));
    }

    private static int CountInvalidations(Mock<IRibbonUI> ribbon) =>
        ribbon.Invocations.Count(item => item.Method.Name == nameof(IRibbonUI.Invalidate));
}
