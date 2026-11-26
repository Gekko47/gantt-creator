using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using ExcelDna.Integration.CustomUI;
using GanttCreator.Core;
using GanttCreator.Core.Logging;
using GanttCreator.Office;
using Moq;
using static GanttCreator.AddIn.Tests.RibbonTestSupport;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// The ribbon's chart and plot controls: the Chart and Plot Size groups, the time-scale and margin
/// dropdowns, the size presets, the plot-date rows and the AUTO toggles.
/// </summary>
/// <remarks>
/// Split out of <see cref="GanttRibbonTests"/> in slice 8. It joins the same collection, so the
/// process-wide boundary and diagnostics singletons stay serialised with the rest of the family.
/// </remarks>
[Collection("diagnostics-service")]
public sealed class GanttRibbonChartTests
{
    [Fact]
    public void Chart_group_declares_start_finish_and_scale_rows()
    {
        // R5.16a: Plot Area and Plot Time Scale are one "Chart" group. Each
        // date is a single row — caption, box, then the Auto checkbox that
        // carries its own label — and the scale is the third row.
        string xml = Ribbon.GetCustomUI(WorkbookRibbonId)!;
        XDocument doc = XDocument.Parse(xml);
        XNamespace ns = NamespaceCustomUI2010;

        XElement group = doc.Descendants(ns + "group")
            .Single(element => element.Attribute("id")?.Value == RibbonControlIds.PlotAreaGroup);
        Assert.Equal("Chart", group.Attribute("label")?.Value);

        var labelIds = group.Descendants(ns + "labelControl")
            .Select(element => element.Attribute("id")?.Value)
            .ToList();
        Assert.Equal(
            ["lblPlotStartDate", "lblPlotFinishDate", "lblPlotTimeScale"],
            labelIds);

        XElement startCheck = group.Descendants(ns + "checkBox")
            .Single(element => element.Attribute("id")?.Value == "chkPlotStartAuto");
        Assert.Equal("Auto", startCheck.Attribute("label")?.Value);
        Assert.Equal("GetPlotStartAuto", startCheck.Attribute("getPressed")?.Value);
        Assert.Equal("OnPlotStartAutoClick", startCheck.Attribute("onAction")?.Value);

        XElement finishCheck = group.Descendants(ns + "checkBox")
            .Single(element => element.Attribute("id")?.Value == "chkPlotFinishAuto");
        Assert.Equal("Auto", finishCheck.Attribute("label")?.Value);
        Assert.Equal("GetPlotFinishAuto", finishCheck.Attribute("getPressed")?.Value);
        Assert.Equal("OnPlotFinishAutoClick", finishCheck.Attribute("onAction")?.Value);

        XElement startBox = group.Descendants(ns + "editBox")
            .Single(element => element.Attribute("id")?.Value == "edtPlotStartDate");
        Assert.Equal("GetPlotStartDate", startBox.Attribute("getText")?.Value);
        Assert.Equal("GetPlotStartDateEnabled", startBox.Attribute("getEnabled")?.Value);
        Assert.Equal("OnPlotStartDateChange", startBox.Attribute("onChange")?.Value);

        // The box is sized for the default display format dd-mmm-yy.
        Assert.Equal("dd-mmm-yy", startBox.Attribute("sizeString")?.Value);

        XElement finishBox = group.Descendants(ns + "editBox")
            .Single(element => element.Attribute("id")?.Value == "edtPlotFinishDate");
        Assert.Equal("GetPlotFinishDate", finishBox.Attribute("getText")?.Value);
        Assert.Equal("GetPlotFinishDateEnabled", finishBox.Attribute("getEnabled")?.Value);
        Assert.Equal("OnPlotFinishDateChange", finishBox.Attribute("onChange")?.Value);
        Assert.Equal("dd-mmm-yy", finishBox.Attribute("sizeString")?.Value);
    }

    [Fact]
    public void Chart_group_declares_the_time_scale_dropdown()
    {
        // R5.2: the lower band's calendar unit is a closed set
        // (month/quarter/week — the year is the fixed upper band,
        // so a Year lower band would duplicate it). R5.12 folds the
        // month band's two label forms into the same dropdown as two
        // Month rows, so one control selects the scale AND the form
        // and there is nothing to grey out: every row the dropdown
        // offers is a pair the catalogue reader accepts.
        // R5.16a: the dropdown lives in the merged Chart group.
        string xml = Ribbon.GetCustomUI(WorkbookRibbonId)!;
        XDocument doc = XDocument.Parse(xml);
        XNamespace ns = NamespaceCustomUI2010;

        XElement group = doc.Descendants(ns + "group")
            .Single(element => element.Attribute("id")?.Value == RibbonControlIds.PlotAreaGroup);
        Assert.Equal("Chart", group.Attribute("label")?.Value);

        // One dropdown, not two: the label form is a row of the scale
        // dropdown, so there is no second control to keep in step.
        XElement dropdown = group.Descendants(ns + "dropDown")
            .Single(element => element.Attribute("id")?.Value == "ddnPlotTimeScale");
        Assert.Equal("GetTimeScale", dropdown.Attribute("getSelectedItemIndex")?.Value);
        Assert.Null(dropdown.Attribute("getSelectedIndex"));
        Assert.Equal("GetEnabled", dropdown.Attribute("getEnabled")?.Value);
        Assert.Equal("OnTimeScaleChange", dropdown.Attribute("onAction")?.Value);
        Assert.Single(group.Descendants(ns + "dropDown"));

        Assert.Equal(
            ["Month (MM)", "Month (MMM)", "Quarter", "Week"],
            dropdown.Descendants(ns + "item").Select(element => element.Attribute("label")?.Value).ToList());

        // The ids are what the commit path parses: the scale name, with the
        // label form suffixed after a hyphen on the two month rows. A
        // display-only label would be refused at commit time, so the id
        // encoding is part of the contract.
        Assert.Equal(
            ["Month-MM", "Month-MMM", "Quarter", "Week"],
            dropdown.Descendants(ns + "item").Select(element => element.Attribute("id")?.Value).ToList());
    }

    [Fact]
    public void PlotSize_group_declares_dimensions_presets_and_margin()
    {
        // R5.2: the plot-size group carries the read-only width/
        // height displays (getText only — they are outputs, never
        // inputs), the four output-size preset buttons, and the
        // margin combobox with its custom-centimetre edit box.
        // The width/height boxes are permanently disabled through
        // the dedicated GetPlotDimensionsEnabled getter (R5.2 D4);
        // the preset buttons are gated on the shared workbook fact;
        // the custom-cm box is gated on the margin being Custom.
        string xml = Ribbon.GetCustomUI(WorkbookRibbonId)!;
        XDocument doc = XDocument.Parse(xml);
        XNamespace ns = NamespaceCustomUI2010;

        XElement group = doc.Descendants(ns + "group")
            .Single(element => element.Attribute("id")?.Value == "grpPlotSize");
        Assert.Equal("Size", group.Attribute("label")?.Value);

        // R5.16a: the width/height readouts are labels, not disabled edit boxes.
        XElement widthLabel = group.Descendants(ns + "labelControl")
            .Single(element => element.Attribute("id")?.Value == "lblPlotWidth");
        Assert.Equal("GetPlotWidthLabel", widthLabel.Attribute("getLabel")?.Value);

        XElement heightLabel = group.Descendants(ns + "labelControl")
            .Single(element => element.Attribute("id")?.Value == "lblPlotHeight");
        Assert.Equal("GetPlotHeightLabel", heightLabel.Attribute("getLabel")?.Value);

        Assert.DoesNotContain(
            group.Descendants(ns + "editBox"),
            element => element.Attribute("id")?.Value is "edtPlotWidth" or "edtPlotHeight");

        // The four preset buttons are one dropdown whose item ids are the
        // durable SizePreset keys.
        XElement sizeDropdown = group.Descendants(ns + "dropDown")
            .Single(element => element.Attribute("id")?.Value == RibbonControlIds.SizeDropdown);
        Assert.Equal("GetSize", sizeDropdown.Attribute("getSelectedItemIndex")?.Value);
        Assert.Equal("OnSizeChange", sizeDropdown.Attribute("onAction")?.Value);
        Assert.Equal(
            ["A4Portrait", "A4Landscape", "Presentation16x9", "Presentation4x3"],
            sizeDropdown.Descendants(ns + "item").Select(element => element.Attribute("id")?.Value));
        Assert.DoesNotContain(
            group.Descendants(ns + "button"),
            element => element.Attribute("onAction")?.Value == "OnPresetClick");

        XElement marginDropdown = group.Descendants(ns + "dropDown")
            .Single(element => element.Attribute("id")?.Value == "ddnMargin");
        Assert.Equal("GetMargin", marginDropdown.Attribute("getSelectedItemIndex")?.Value);
        Assert.Null(marginDropdown.Attribute("getSelectedIndex"));
        Assert.Equal("GetEnabled", marginDropdown.Attribute("getEnabled")?.Value);
        Assert.Equal("OnMarginChange", marginDropdown.Attribute("onAction")?.Value);
        var marginLabels = marginDropdown.Descendants(ns + "item")
            .Select(element => element.Attribute("label")?.Value)
            .ToList();
        Assert.Equal(
            ["Narrow (0.635 cm)", "Normal (1.78 cm)", "Wide (2.54 cm)", "Custom"],
            marginLabels);

        XElement marginCmBox = group.Descendants(ns + "editBox")
            .Single(element => element.Attribute("id")?.Value == "edtMarginCm");
        Assert.Equal("GetMarginCm", marginCmBox.Attribute("getText")?.Value);
        Assert.Equal("GetMarginCmEnabled", marginCmBox.Attribute("getEnabled")?.Value);
        Assert.Equal("OnMarginCmChange", marginCmBox.Attribute("onChange")?.Value);
    }

    /// <summary>
    /// The time-scale commit path forwards the selected item id: an exact
    /// ordinal match persists the scale, so the dropdown and the stored setting
    /// cannot disagree. The routing itself needs no workbook.
    /// </summary>
    [Fact]
    public void OnTimeScaleChange_forwards_the_selected_scale_to_the_service()
    {
        var service = new RibbonStateService();
        service.SetCatalogueWriter(null);

        GanttRibbon.OnTimeScaleChange(null, "Week", service);

        Assert.Equal(GanttTimeScale.Week, service.GetTimeScale());
    }

    /// <summary>
    /// The two Month rows commit the scale AND the label form together, so one
    /// selection is what reaches storage: the dropdown and the stored setting
    /// cannot disagree on whether the band shows 01 or Jan.
    /// </summary>
    [Fact]
    public void OnTimeScaleChange_commits_the_month_label_form_with_the_scale()
    {
        var service = new RibbonStateService();
        service.SetCatalogueWriter(null);

        GanttRibbon.OnTimeScaleChange(null, "Month-MM", service);

        Assert.Equal(GanttTimeScale.Month, service.GetTimeScale());
        Assert.Equal(GanttPeriodLabelFormat.MM, service.GetPeriodLabelFormat());

        GanttRibbon.OnTimeScaleChange(null, "Month-MMM", service);

        Assert.Equal(GanttTimeScale.Month, service.GetTimeScale());
        Assert.Equal(GanttPeriodLabelFormat.MMM, service.GetPeriodLabelFormat());
    }

    /// <summary>
    /// A suffix-less id names the scale only, so the service supplies the
    /// scale's canonical form. Quarter and Week therefore land on their own
    /// format even when a month form was stored a moment earlier.
    /// </summary>
    [Fact]
    public void OnTimeScaleChange_normalises_the_format_for_a_suffix_less_scale()
    {
        var service = new RibbonStateService();
        service.SetCatalogueWriter(null);
        GanttRibbon.OnTimeScaleChange(null, "Month-MM", service);

        GanttRibbon.OnTimeScaleChange(null, "Quarter", service);

        Assert.Equal(GanttTimeScale.Quarter, service.GetTimeScale());
        Assert.Equal(GanttPeriodLabelFormat.Quarter, service.GetPeriodLabelFormat());
    }

    /// <summary>
    /// An id the closed-set parsers refuse leaves the stored pair untouched: an
    /// unknown scale, an unknown format, and a format the named scale cannot
    /// carry are all refused rather than defaulted, so a mistyped or stale
    /// selection cannot silently redraw the period band.
    /// </summary>
    [Fact]
    public void OnTimeScaleChange_refuses_an_unusable_selection()
    {
        var service = new RibbonStateService();
        service.SetCatalogueWriter(null);
        GanttRibbon.OnTimeScaleChange(null, "Month-MMM", service);

        GanttRibbon.OnTimeScaleChange(null, "Year", service);
        GanttRibbon.OnTimeScaleChange(null, "Month-mmm", service);
        GanttRibbon.OnTimeScaleChange(null, "Month-Quarter", service);
        GanttRibbon.OnTimeScaleChange(null, "Quarter-MM", service);
        GanttRibbon.OnTimeScaleChange(null, string.Empty, service);

        Assert.Equal(GanttTimeScale.Month, service.GetTimeScale());
        Assert.Equal(GanttPeriodLabelFormat.MMM, service.GetPeriodLabelFormat());
    }

    /// <summary>
    /// The getter's selected index is derived from the PAIR, and it is the
    /// exact inverse of the id parse: each of the four rows round-trips, so the
    /// highlighted row can never disagree with the committed setting. Without
    /// this the getter would return the bare scale ordinal and the Month rows
    /// would both highlight the first one.
    /// </summary>
    [Fact]
    public void GetTimeScale_maps_each_selection_to_its_own_row()
    {
        var service = new RibbonStateService();
        service.SetCatalogueWriter(null);

        GanttRibbon.OnTimeScaleChange(null, "Month-MM", service);
        Assert.Equal(0, GanttRibbon.GetTimeScale(null, service));

        GanttRibbon.OnTimeScaleChange(null, "Month-MMM", service);
        Assert.Equal(1, GanttRibbon.GetTimeScale(null, service));

        GanttRibbon.OnTimeScaleChange(null, "Quarter", service);
        Assert.Equal(2, GanttRibbon.GetTimeScale(null, service));

        GanttRibbon.OnTimeScaleChange(null, "Week", service);
        Assert.Equal(3, GanttRibbon.GetTimeScale(null, service));
    }

    /// <summary>
    /// The margin commit path forwards the selected item id
    /// through the closed-set parser: an exact ordinal match
    /// persists the preset, so the combobox and the stored
    /// setting cannot disagree.
    /// </summary>
    [Fact]
    public void OnMarginChange_forwards_the_selected_margin_to_the_service()
    {
        var service = new RibbonStateService();
        service.SetCatalogueWriter(null);

        GanttRibbon.OnMarginChange(null, "Custom", service);

        Assert.Equal(GanttPlotMargin.Custom, service.GetMargin());
    }

    /// <summary>
    /// An unknown item id is refused by the closed-set parser,
    /// so the stored margin preset is untouched.
    /// </summary>
    [Fact]
    public void OnMarginChange_refuses_an_unknown_margin_id()
    {
        var service = new RibbonStateService();
        service.SetCatalogueWriter(null);
        service.SetMargin(GanttPlotMargin.Normal);

        GanttRibbon.OnMarginChange(null, "Narrow-ish", service);

        Assert.Equal(GanttPlotMargin.Normal, service.GetMargin());
    }

    /// <summary>
    /// The custom-margin commit path parses the edited text
    /// with the invariant culture and forwards the centimetre
    /// value, so a locale-independent commit persists the
    /// typed margin.
    /// </summary>
    [Fact]
    public void OnMarginCmChange_forwards_the_parsed_cm_to_the_service()
    {
        var service = new RibbonStateService();
        service.SetCatalogueWriter(null);

        GanttRibbon.OnMarginCmChange(null, "2.5", service);

        Assert.Equal(2.5, service.GetMarginCm());
    }

    /// <summary>
    /// An unparsable commit is a no-op, so a partial keystroke
    /// cannot corrupt the stored margin: the previous value
    /// survives.
    /// </summary>
    [Fact]
    public void OnMarginCmChange_refuses_an_unparsable_value()
    {
        var service = new RibbonStateService();
        service.SetCatalogueWriter(null);
        service.SetMarginCm(1.0);

        GanttRibbon.OnMarginCmChange(null, "abc", service);

        Assert.Equal(1.0, service.GetMarginCm());
    }

    /// <summary>
    /// Excel's <c>dropDown.onAction</c> contract passes the control, the chosen
    /// item id and its index. The item id selects the preset, so choosing A4
    /// landscape stores it and the width/height displays and the next Refresh
    /// agree. The control's own id is deliberately a different value here: the
    /// contract is the item id, and reading <c>control.Id</c> was the live
    /// defect this test now prevents.
    /// </summary>
    [Fact]
    public void OnSizeChange_routes_the_selected_item_id_to_the_service()
    {
        var service = new RibbonStateService();
        service.SetCatalogueWriter(null);
        var control = new Mock<IRibbonControl>();
        control.SetupGet(c => c.Id).Returns(RibbonControlIds.SizeDropdown);

        GanttRibbon.OnSizeChange(control.Object, "A4Landscape", service);

        Assert.Equal(SizePresets.A4Landscape, service.GetPreset());
    }

    /// <summary>
    /// An unknown item id is refused — never defaulted to a preset —
    /// so a mistyped or renamed item cannot silently render the chart on
    /// the wrong paper: the stored preset is untouched.
    /// </summary>
    [Fact]
    public void OnSizeChange_refuses_an_unknown_item_id()
    {
        var service = new RibbonStateService();
        service.SetCatalogueWriter(null);
        service.SetPreset(SizePresets.A4Landscape);
        var control = new Mock<IRibbonControl>();
        control.SetupGet(c => c.Id).Returns(RibbonControlIds.SizeDropdown);

        GanttRibbon.OnSizeChange(control.Object, "A3", service);

        Assert.Equal(SizePresets.A4Landscape, service.GetPreset());
    }

    /// <summary>
    /// The preset pressed-state getter marks the active preset:
    /// the control whose id names the stored preset reports
    /// pressed, and every other preset button reports not
    /// pressed, so exactly one preset reads as active at a
    /// time. The comparison is by durable key, so a preset
    /// record compared by reference would misreport.
    /// </summary>
    [Fact]
    public void GetSize_projects_the_stored_preset_onto_the_item_list()
    {
        // The item list mirrors SizePresets.All, so the selected index is a
        // projection of the stored preset rather than a second list.
        var service = new RibbonStateService();
        service.SetCatalogueWriter(null);

        service.SetPreset(SizePresets.Presentation16x9);
        Assert.Equal(2, GanttRibbon.GetSize(null, service));

        // A workbook that names no preset reports the first item, because
        // Excel has no "nothing selected" state for this callback.
        var empty = new RibbonStateService();
        empty.SetCatalogueWriter(null);
        Assert.Equal(0, GanttRibbon.GetSize(null, empty));
    }

    /// <summary>
    /// The live defect of 2026-11-13: on a blank sheet <c>GetSize</c> returned
    /// -1, and Excel rejected it with "The callback function GetSize on control
    /// ddnSize returned an invalid value". Every reachable state must name a real
    /// item, including the state where no workbook, table, or preset exists.
    /// </summary>
    [Fact]
    public void GetSize_never_returns_an_negative_index()
    {
        var unarmed = new RibbonStateService();
        unarmed.SetCatalogueWriter(null);

        int index = GanttRibbon.GetSize(null, unarmed);

        Assert.InRange(index, 0, SizePresets.All.Count - 1);
    }

    /// <summary>
    /// The dropdown's items <em>are</em> the catalogue, in order. An item added
    /// to one without the other would let <c>GetSize</c> name an index the
    /// dropdown cannot resolve — the same "invalid value" the host reported for
    /// a negative index.
    /// </summary>
    [Fact]
    public void The_size_dropdown_items_are_the_size_preset_catalogue_in_order()
    {
        XDocument document = XDocument.Parse(Ribbon.GetCustomUI(WorkbookRibbonId)!);
        XNamespace ns = document.Root!.Name.Namespace;
        string[] itemIds =
        [
            .. document.Descendants(ns + "dropDown")
                .Where(dropDown => (string?)dropDown.Attribute("id") == RibbonControlIds.SizeDropdown)
                .SelectMany(dropDown => dropDown.Elements(ns + "item"))
                .Select(item => (string?)item.Attribute("id") ?? string.Empty),
        ];

        Assert.Equal(SizePresets.All.Select(preset => preset.Key), itemIds);
    }

    /// <summary>
    /// The edit-box commit path forwards the edited text: an exact dd/MM/yyyy
    /// commit persists the normalised date; the routing itself needs no workbook.
    /// </summary>
    [Fact]
    public void OnPlotStartDateChange_forwards_the_edited_text_to_the_service()
    {
        var service = new RibbonStateService();
        service.SetCatalogueWriter(null);

        GanttRibbon.OnPlotStartDateChange(null, "15-Mar-26", service);

        Assert.Equal("15-Mar-26", service.GetPlotStartDate());
        Assert.False(service.IsPlotStartAuto());
    }

    /// <summary>
    /// The setter re-invalidates exactly its own row — the date box, the calendar
    /// button and the Auto box — and never the whole tab, so the clicked control
    /// repaints from the new snapshot at the cost of three getter queries instead
    /// of every getter on the ribbon (R5.1a C3).
    /// </summary>
    [Fact]
    public void SetPlotStartAuto_invalidates_only_the_start_row()
    {
        var ribbon = new Mock<IRibbonUI>();
        var service = new RibbonStateService();
        service.SetCatalogueWriter(null);
        service.SetRibbon(ribbon.Object);

        service.SetPlotStartAuto(true);

        Assert.Equal(
            [
                RibbonControlIds.PlotStartDate,
                RibbonControlIds.PlotStartCalendar,
                RibbonControlIds.PlotStartAuto,
            ],
            RibbonInvalidatedIds(ribbon));
        Assert.Equal(
            0,
            ribbon.Invocations.Count(invocation => invocation.Method.Name == nameof(IRibbonUI.Invalidate)));
    }

    /// <summary>The finish-end analogue: one commit, three row controls, no blanket invalidate.</summary>
    [Fact]
    public void SetPlotFinishAuto_invalidates_only_the_finish_row()
    {
        var ribbon = new Mock<IRibbonUI>();
        var service = new RibbonStateService();
        service.SetCatalogueWriter(null);
        service.SetRibbon(ribbon.Object);

        service.SetPlotFinishAuto(true);

        Assert.Equal(
            [
                RibbonControlIds.PlotFinishDate,
                RibbonControlIds.PlotFinishCalendar,
                RibbonControlIds.PlotFinishAuto,
            ],
            RibbonInvalidatedIds(ribbon));
        Assert.Equal(
            0,
            ribbon.Invocations.Count(invocation => invocation.Method.Name == nameof(IRibbonUI.Invalidate)));
    }
    /// <summary>
    /// Positive for the true-bounds display: with both ends AUTO and event
    /// data present, the boxes show the month-snapped chart bounds Refresh
    /// renders (01-Mar-26–30-Apr-26 for 10 Mar–31 Mar spans with the
    /// default 3-day pad), not the raw padded data extent (07/03–03/04).
    /// </summary>
    [Fact]
    public void Plot_date_getters_show_the_true_chart_bounds_in_auto_mode()
    {
        var service = new RibbonStateService();
        service.SetCatalogueWriter(null);
        var catalogue = new Mock<IConfigCatalogueReader>();
        _ = catalogue.Setup(r => r.Read()).Returns(ConfigReadOutcome.Ok(
            "wb",
            new Dictionary<string, string>
            {
                ["PlotStartMode"] = "DataRange",
                ["PlotFinishMode"] = "DataRange",
                ["PlotStartDate"] = string.Empty,
                ["PlotFinishDate"] = string.Empty,
            },
            GanttStyleRegistry.Empty, new Dictionary<string, double>()));
        service.SetCatalogueReader(catalogue.Object);
        var table = new Mock<IGanttTableReader>();
        _ = table.Setup(r => r.Read()).Returns(GanttTableReadOutcome.Ok(
        [
            PlotBoundsRow(1, "10/03/2026", "31/03/2026"),
        ]));
        service.SetTableReader(table.Object);

        service.Refresh();

        Assert.Equal("01-Mar-26", service.GetPlotStartDate());
        Assert.Equal("30-Apr-26", service.GetPlotFinishDate());
    }

    [Fact]
    public void OnPlotFinishDateChange_forwards_the_edited_text_to_the_service()
    {
        var service = new RibbonStateService();
        service.SetCatalogueWriter(null);

        GanttRibbon.OnPlotFinishDateChange(null, "30-Jun-26", service);

        Assert.Equal("30-Jun-26", service.GetPlotFinishDate());
        Assert.False(service.IsPlotFinishAuto());
    }

    [Fact]
    public void Plot_date_getters_answer_from_the_snapshot()
    {
        var adapter = new Mock<IExcelApplicationAdapter>();
        _ = adapter.Setup(a => a.HasActiveWorkbook()).Returns(true);
        var tableReader = new Mock<IGanttTableReader>();
        _ = tableReader.Setup(t => t.Read()).Returns(GanttTableReadOutcome.Ok([]));
        var service = new RibbonStateService();
        service.SetApplicationAdapter(adapter.Object);
        service.SetTableReader(tableReader.Object);
        service.SetCatalogueWriter(null);
        service.SetPlotStartDate("15-Mar-26");
        service.SetPlotFinishDate("30-Jun-26");
        service.Refresh();

        var control = new Mock<IRibbonControl>();

        Assert.Equal("15-Mar-26", GanttRibbon.GetPlotStartDate(control.Object, service));
        Assert.Equal("30-Jun-26", GanttRibbon.GetPlotFinishDate(control.Object, service));

        // Explicit on both ends: the edit boxes must be enabled, and the
        // AUTO checkboxes unchecked.
        Assert.True(GanttRibbon.GetPlotStartDateEnabled(control.Object, service));
        Assert.True(GanttRibbon.GetPlotFinishDateEnabled(control.Object, service));
        Assert.False(GanttRibbon.GetPlotStartAuto(control.Object, service));
        Assert.False(GanttRibbon.GetPlotFinishAuto(control.Object, service));
    }

    [Fact]
    public void Plot_date_getters_disable_the_edit_box_while_auto_is_on()
    {
        var service = new RibbonStateService();

        var control = new Mock<IRibbonControl>();

        Assert.False(GanttRibbon.GetPlotStartDateEnabled(control.Object, service));
        Assert.False(GanttRibbon.GetPlotFinishDateEnabled(control.Object, service));
        Assert.True(GanttRibbon.GetPlotStartAuto(control.Object, service));
        Assert.True(GanttRibbon.GetPlotFinishAuto(control.Object, service));
    }

    [Fact]
    public void Plot_date_getText_degrades_to_empty_when_the_snapshot_probe_fails()
    {
        // Fail-open is the contract: a getter that throws into Excel's
        // dispatch breaks the whole ribbon, so every plot getter degrades to
        // the safe direction — empty text, AUTO checked (the initial state),
        // and the edit box enabled rather than permanently grey.
        var broken = new ThrowingStateService();
        var control = new Mock<IRibbonControl>();

        Assert.Equal(string.Empty, GanttRibbon.GetPlotStartDate(control.Object, broken));
        Assert.Equal(string.Empty, GanttRibbon.GetPlotFinishDate(control.Object, broken));
        Assert.True(GanttRibbon.GetPlotStartAuto(control.Object, broken));
        Assert.True(GanttRibbon.GetPlotFinishAuto(control.Object, broken));
        Assert.True(GanttRibbon.GetPlotStartDateEnabled(control.Object, broken));
        Assert.True(GanttRibbon.GetPlotFinishDateEnabled(control.Object, broken));
    }

    private sealed class ThrowingStateService : RibbonStateService
    {
        internal override string GetPlotStartDate() => throw new InvalidOperationException("probe failed");

        internal override string GetPlotFinishDate() => throw new InvalidOperationException("probe failed");

        internal override bool IsPlotStartAuto() => throw new InvalidOperationException("probe failed");

        internal override bool IsPlotFinishAuto() => throw new InvalidOperationException("probe failed");

        internal override bool GetEnabled(string? controlId) => true;
    }

    /// <summary>Null dependencies are refused at the boundary, not at first use.</summary>
    [Fact]
    public void OnPlotStartAutoClick_throws_for_a_null_state_service()
    {
        var control = new Mock<IRibbonControl>();

        Assert.Throws<ArgumentNullException>(
            () => GanttRibbon.OnPlotStartAutoClick(control.Object, true, null!));
    }

    /// <summary>Null dependencies are refused at the boundary, not at first use.</summary>
    [Fact]
    public void OnPlotFinishAutoClick_throws_for_a_null_state_service()
    {
        var control = new Mock<IRibbonControl>();

        Assert.Throws<ArgumentNullException>(
            () => GanttRibbon.OnPlotFinishAutoClick(control.Object, true, null!));
    }

    /// <summary>
    /// The checkbox honours Excel's pressed state (fix plan ruling 2): checking
    /// AUTO enters automatic mode, unchecking arms explicit mode.
    /// </summary>
    [Fact]
    public void OnPlotStartAutoClick_honours_the_pressed_state_Excel_reports()
    {
        var control = new Mock<IRibbonControl>();

        var autoService = new RibbonStateService();
        autoService.SetCatalogueWriter(null);
        GanttRibbon.OnPlotStartAutoClick(control.Object, true, autoService);
        Assert.True(autoService.IsPlotStartAuto());

        var explicitService = new RibbonStateService();
        explicitService.SetCatalogueWriter(null);
        explicitService.SetPlotStartDate("15-Mar-26");
        GanttRibbon.OnPlotStartAutoClick(control.Object, false, explicitService);
        Assert.False(explicitService.IsPlotStartAuto());
    }

    /// <summary>
    /// The checkbox honours Excel's pressed state (fix plan ruling 2): checking
    /// AUTO enters automatic mode, unchecking arms explicit mode.
    /// </summary>
    [Fact]
    public void OnPlotFinishAutoClick_honours_the_pressed_state_Excel_reports()
    {
        var control = new Mock<IRibbonControl>();

        var autoService = new RibbonStateService();
        autoService.SetCatalogueWriter(null);
        GanttRibbon.OnPlotFinishAutoClick(control.Object, true, autoService);
        Assert.True(autoService.IsPlotFinishAuto());

        var explicitService = new RibbonStateService();
        explicitService.SetCatalogueWriter(null);
        explicitService.SetPlotFinishDate("30-Jun-26");
        GanttRibbon.OnPlotFinishAutoClick(control.Object, false, explicitService);
        Assert.False(explicitService.IsPlotFinishAuto());
    }

    /// <summary>
    /// Positive for the uncheck fix: unchecking AUTO with a stored date
    /// persists Explicit with that date, so the next workbook load cannot
    /// revert the box to checked. The date commit's own write is not counted:
    /// only the uncheck's persist is asserted.
    /// </summary>
    [Fact]
    public void SetPlotStartAuto_uncheck_with_a_stored_date_persists_explicit()
    {
        // CommandBoundary.Instance is a process-wide singleton: save
        // its log and restore it afterwards so this test's mock does
        // not leak into other tests. SetLog(null) throws, so a null
        // original is restored by resetting the boundary to its
        // pristine state.
        var originalLog = CommandBoundary.Instance.GetLog();
        var log = new Mock<IRollingLog>();
        CommandBoundary.Instance.SetLog(log.Object);
        try
        {
            var service = new RibbonStateService();
            service.SetCatalogueWriter(null);
            service.SetPlotStartDate("15-Mar-26");
            var writer = new Mock<IConfigCatalogueWriter>(MockBehavior.Strict);
            _ = writer
                .Setup(w => w.WriteSettings(It.Is<IReadOnlyDictionary<string, string>>(
                    settings =>
                        settings["PlotStartMode"] == "Explicit"
                        && settings["PlotStartDate"] == "15-Mar-26")))
                .Returns(ConfigWriteOutcome.Ok());
            service.SetCatalogueWriter(writer.Object);

            service.SetPlotStartAuto(false);

            writer.Verify(
                w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()),
                Times.Once);
            Assert.False(service.IsPlotStartAuto());
            Assert.Equal("15-Mar-26", service.GetPlotStartDate());
        }
        finally
        {
            if (originalLog is null)
            {
                CommandBoundary.Reset();
            }
            else
            {
                CommandBoundary.Instance.SetLog(originalLog);
            }
        }
    }

    /// <summary>
    /// Positive for the uncheck fix on the finish end: unchecking AUTO with a
    /// stored date persists Explicit with that date. The date commit's own
    /// write is not counted: only the uncheck's persist is asserted.
    /// </summary>
    [Fact]
    public void SetPlotFinishAuto_uncheck_with_a_stored_date_persists_explicit()
    {
        // CommandBoundary.Instance is a process-wide singleton: save
        // its log and restore it afterwards so this test's mock does
        // not leak into other tests. SetLog(null) throws, so a null
        // original is restored by resetting the boundary to its
        // pristine state.
        var originalLog = CommandBoundary.Instance.GetLog();
        var log = new Mock<IRollingLog>();
        CommandBoundary.Instance.SetLog(log.Object);
        try
        {
            var service = new RibbonStateService();
            service.SetCatalogueWriter(null);
            service.SetPlotFinishDate("30-Jun-26");
            var writer = new Mock<IConfigCatalogueWriter>(MockBehavior.Strict);
            _ = writer
                .Setup(w => w.WriteSettings(It.Is<IReadOnlyDictionary<string, string>>(
                    settings =>
                        settings["PlotFinishMode"] == "Explicit"
                        && settings["PlotFinishDate"] == "30-Jun-26")))
                .Returns(ConfigWriteOutcome.Ok());
            service.SetCatalogueWriter(writer.Object);

            service.SetPlotFinishAuto(false);

            writer.Verify(
                w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()),
                Times.Once);
            Assert.False(service.IsPlotFinishAuto());
            Assert.Equal("30-Jun-26", service.GetPlotFinishDate());
        }
        finally
        {
            if (originalLog is null)
            {
                CommandBoundary.Reset();
            }
            else
            {
                CommandBoundary.Instance.SetLog(originalLog);
            }
        }
    }

    /// <summary>Null dependencies are refused at the boundary, not at first use.</summary>
    [Fact]
    public void OnPlotStartDateChange_throws_for_a_null_state_service()
    {
        var control = new Mock<IRibbonControl>();

        Assert.Throws<ArgumentNullException>(
            () => GanttRibbon.OnPlotStartDateChange(control.Object, "15-Mar-26", null!));
    }

    /// <summary>Null dependencies are refused at the boundary, not at first use.</summary>
    [Fact]
    public void OnPlotFinishDateChange_throws_for_a_null_state_service()
    {
        var control = new Mock<IRibbonControl>();

        Assert.Throws<ArgumentNullException>(
            () => GanttRibbon.OnPlotFinishDateChange(control.Object, "30-Jun-26", null!));
    }

    /// <summary>
    /// Null dependencies are refused at the boundary, not at first use.
    /// </summary>
    [Fact]
    public void GetPlotStartDate_throws_for_a_null_state_service()
    {
        var control = new Mock<IRibbonControl>();

        Assert.Throws<ArgumentNullException>(
            () => GanttRibbon.GetPlotStartDate(control.Object, null!));
    }
}
