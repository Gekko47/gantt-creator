using System.Reflection;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using ExcelDna.Integration.CustomUI;
using GanttCreator.Core;
using static GanttCreator.AddIn.Tests.RibbonTestSupport;

// CA2000: test fixtures are deliberately not disposed — the code under test
// (GanttRibbon) does not own disposable resources, and these are plain strings
// / reflection results. Scoped to this file.
#pragma warning disable CA2000

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// The RibbonX surface contract for <see cref="GanttRibbon"/>, read from the XML
/// <see cref="GanttRibbon.GetCustomUI"/> returns: the tab and group shape, every callback that must
/// resolve, the gated and labelled controls, the glyph map, the entity menus and the export group.
/// </summary>
/// <remarks>
/// These facts pin the document rather than a click. None of them require Excel, and none of them
/// drive a product callback, so a failure here names the ribbon markup or its callback contract
/// rather than a command.
/// </remarks>
[Collection("diagnostics-service")]
public sealed class GanttRibbonXmlTests
{
    [Fact]
    public void GanttRibbon_is_public_ComVisible_ExcelRibbon_derived_and_has_parameterless_ctor()
    {
        Type type = typeof(GanttRibbon);

        Assert.True(type.IsPublic, "Excel-DNA requires a public ribbon type.");
        ComVisibleAttribute? ComVisibleAttribute =
            type.GetCustomAttribute<ComVisibleAttribute>(inherit: true);
        Assert.True(
            ComVisibleAttribute is not null && ComVisibleAttribute.Value == true,
            "The ribbon type must be [ComVisible(true)] for COM registration.");
        Assert.True(
            typeof(ExcelRibbon).IsAssignableFrom(type),
            "Excel-DNA only auto-registers ExcelRibbon descendants (IsRibbonType).");
        System.Reflection.ConstructorInfo? ctor = type.GetConstructor(Type.EmptyTypes);
        Assert.NotNull(ctor);
        Assert.True(
            ctor!.IsPublic,
            "Excel-DNA instantiates the ribbon via Activator.CreateInstance, which requires a public parameterless constructor.");
    }

    [Fact]
    public void GetCustomUI_returns_the_ribbon_xml_for_the_workbook_id()
    {
        string? xml = Ribbon.GetCustomUI(WorkbookRibbonId);

        Assert.NotNull(xml);
        Assert.NotEmpty(xml);
        Assert.Contains("customUI", xml, StringComparison.Ordinal);
    }
    [Fact]
    public void Every_ribbon_group_declares_the_visible_callback()
    {
        Assert.Empty(GroupsMissingVisibleCallback(Ribbon.GetCustomUI(WorkbookRibbonId)!));
    }

    [Fact]
    public void The_visible_callback_validator_flags_a_group_that_omits_it()
    {
        // Positive test: a group without getVisible must be named, not silently
        // accepted. Without this, the check above would pass by matching nothing.
        const string xml = """
            <customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
              <ribbon><tabs><tab id="tab">
                <group id="grpGood" getVisible="GetVisible"/>
                <group id="grpBad"/>
              </tab></tabs></ribbon>
            </customUI>
            """;

        Assert.Equal(["grpBad"], GroupsMissingVisibleCallback(xml));
    }

    [Fact]
    public void Get_started_holds_exactly_the_large_initialise_button()
    {
        XDocument doc = XDocument.Parse(Ribbon.GetCustomUI(WorkbookRibbonId)!);
        XNamespace ns = doc.Root!.Name.Namespace;
        XElement group = Assert.Single(
            doc.Descendants(ns + "group"),
            element => (string?)element.Attribute("id") == RibbonControlIds.GetStartedGroup);

        XElement button = Assert.Single(group.Elements());
        Assert.Equal(RibbonControlIds.InitialiseSheet, (string?)button.Attribute("id"));
        Assert.Equal("large", (string?)button.Attribute("size"));
    }

    [Fact]
    public void The_custom_style_selector_lives_in_the_style_group()
    {
        // Owner ruling 2026-11-17: a row's style is a Style subject, not a Labels one.
        // The selector heads the Style group, so it reads "Custom style [dropdown] |
        // Theme settings... | Chart title...", and the Labels group keeps only the two
        // label-position dropdowns. Pinned so a later layout pass cannot move it back
        // without this fact, its caption and the two entity supertips following.
        XDocument doc = XDocument.Parse(Ribbon.GetCustomUI(WorkbookRibbonId)!);
        XNamespace ns = doc.Root!.Name.Namespace;

        XElement style = Assert.Single(
            doc.Descendants(ns + "group"),
            element => (string?)element.Attribute("id") == RibbonControlIds.StyleThemeGroup);
        XElement labels = Assert.Single(
            doc.Descendants(ns + "group"),
            element => (string?)element.Attribute("id") == RibbonControlIds.LabelsGroup);

        Assert.Equal("boxCustomStyle", (string?)style.Elements().First().Attribute("id"));

        Assert.Single(
            style.Descendants(ns + "dropDown"),
            element => (string?)element.Attribute("id") == RibbonControlIds.CustomStyle);
        XElement caption = Assert.Single(
            style.Descendants(ns + "labelControl"),
            element => (string?)element.Attribute("id") == "lblCustomStyle");
        Assert.Equal("Custom style", (string?)caption.Attribute("label"));

        Assert.DoesNotContain(
            labels.Descendants(ns + "dropDown"),
            element => (string?)element.Attribute("id") == RibbonControlIds.CustomStyle);
    }

    [Fact]
    public void The_sheet_group_puts_a_large_refresh_beside_one_stacked_column()
    {
        // R5.18e (owner ruling 2026-11-13): Refresh is the group's primary action, so it
        // is large and the three secondary controls stack beside it in one vertical box,
        // in the owner's order: warnings, cells, chart.
        // R5.20 (owner ruling 2026-11-16): the primary is now a splitButton so the automatic
        // refresh switch lives in its menu, and `size="large"` belongs on the splitButton —
        // every other split button in this document puts it there and nowhere else.
        XDocument doc = XDocument.Parse(Ribbon.GetCustomUI(WorkbookRibbonId)!);
        XNamespace ns = doc.Root!.Name.Namespace;
        XElement group = Assert.Single(
            doc.Descendants(ns + "group"),
            element => (string?)element.Attribute("id") == RibbonControlIds.SheetGroup);

        XElement[] children = [.. group.Elements()];
        Assert.Equal(2, children.Length);

        XElement split = children[0];
        Assert.Equal("splitButton", split.Name.LocalName);
        Assert.Equal(RibbonControlIds.RefreshSplit, (string?)split.Attribute("id"));
        Assert.Equal("large", (string?)split.Attribute("size"));

        XElement refresh = Assert.Single(split.Elements(ns + "button"));
        Assert.Equal(RibbonControlIds.RefreshSheet, (string?)refresh.Attribute("id"));
        Assert.Null(refresh.Attribute("size"));

        XElement menu = Assert.Single(split.Elements(ns + "menu"));
        Assert.Equal(RibbonControlIds.RefreshMenu, (string?)menu.Attribute("id"));
        XElement toggle = Assert.Single(menu.Elements(ns + "toggleButton"));
        Assert.Equal(RibbonControlIds.ToggleAutoRefresh, (string?)toggle.Attribute("id"));

        XElement stack = children[1];
        Assert.Equal("box", stack.Name.LocalName);
        Assert.Equal("vertical", (string?)stack.Attribute("boxStyle"));
        string?[] stacked = [.. stack.Elements().Select(element => (string?)element.Attribute("id"))];
        Assert.Equal(
            new string?[] { RibbonControlIds.Warnings, RibbonControlIds.ToggleCellLock, RibbonControlIds.ToggleChartLock },
            stacked);
    }

    [Fact]
    public void Both_lock_toggles_state_their_current_state_through_getLabel()
    {
        // R5.18f (owner ruling 2026-11-13): the label describes the state the toggle is in, because
        // the pressed state already says what is true. The icon is deliberately unchanged: no
        // open-padlock imageMso exists in the gallery and the idMso table, so an action-shaped
        // label would contradict the glyph as well as the toggle.
        XDocument doc = XDocument.Parse(Ribbon.GetCustomUI(WorkbookRibbonId)!);
        XNamespace ns = doc.Root!.Name.Namespace;

        foreach ((string id, string callback) in new[]
        {
            (RibbonControlIds.ToggleCellLock, nameof(GanttRibbon.GetToggleCellLockLabel)),
            (RibbonControlIds.ToggleChartLock, nameof(GanttRibbon.GetToggleChartLockLabel)),
            // R5.20: the automatic-refresh toggle follows the same rule. Its state is a
            // persisted setting rather than a protection flag, and the wording is owned by
            // RibbonStateService.AutoRefreshLabel so the three toggles read alike.
            (RibbonControlIds.ToggleAutoRefresh, nameof(GanttRibbon.GetToggleAutoRefreshLabel)),
        })
        {
            XElement toggle = Assert.Single(
                doc.Descendants(ns + "toggleButton"),
                element => (string?)element.Attribute("id") == id);
            Assert.Null(toggle.Attribute("label"));
            Assert.Equal(callback, (string?)toggle.Attribute("getLabel"));
        }
    }

    [Theory]
    [InlineData(true, "Chart locked")]
    [InlineData(false, "Chart unlocked")]
    public void The_chart_lock_label_states_the_current_state(bool locked, string expected) =>
        Assert.Equal(expected, RibbonStateService.ChartLockLabel(locked));

    [Theory]
    [InlineData(true, "Cells locked")]
    [InlineData(false, "Cells unlocked")]
    public void The_cell_lock_label_states_the_current_state(bool locked, string expected) =>
        Assert.Equal(expected, RibbonStateService.CellLockLabel(locked));

    [Fact]
    public void Ribbon_xml_parses_and_uses_the_office_2009_namespace()
    {
        string xml = Ribbon.GetCustomUI(WorkbookRibbonId)!;

        XDocument doc = XDocument.Parse(xml);
        XElement? root = doc.Root;
        Assert.NotNull(root);
        Assert.Equal("customUI", root.Name.LocalName, ignoreCase: false);

        // The 2009/07 namespace targets Excel 2010+, the supported baseline.
        string ns = root.Name.NamespaceName;
        Assert.Equal(NamespaceCustomUI2010, ns);
    }

    [Fact]
    public void Ribbon_declares_exactly_one_gantt_creator_tab_with_expected_id_and_label()
    {
        string xml = Ribbon.GetCustomUI(WorkbookRibbonId)!;
        XDocument doc = XDocument.Parse(xml);

        // Elements live in the CustomUI 2009/07 default namespace, so queries
        // must be namespace-qualified to match.
        XNamespace ns = NamespaceCustomUI2010;
        List<XElement> tabs = doc.Descendants(ns + "tab").ToList();
        Assert.Single(tabs);

        XElement tab = tabs[0];
        Assert.Equal(
            "tabGanttCreator",
            tab.Attribute("id")?.Value,
            ignoreCase: false);
        Assert.Equal(
            "Gantt Creator",
            tab.Attribute("label")?.Value,
            ignoreCase: false);
    }

    [Fact]
    public void Callback_contract_all_callbacks_in_the_real_xml_resolve_to_public_methods()
    {
        string xml = Ribbon.GetCustomUI(WorkbookRibbonId)!;

        List<string> unresolved = ValidateCallbacks(xml, typeof(GanttRibbon));

        Assert.Empty(unresolved);
    }

    [Fact]
    public void Callback_contract_reports_a_missing_callback_method_positive_test()
    {
        // A deliberately broken document whose onAction references a method that
        // does not exist on the ribbon type. Proves the validator actually
        // catches violations rather than vacuously passing.
        const string brokenXml = @"<?xml version='1.0' encoding='utf-8'?>
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab id='tabGanttCreator' label='Gantt Creator'>
        <group id='grpPlaceholder' label='Gantt Creator'>
          <button id='btnBroken' label='Broken' onAction='DoesNotExist'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

        List<string> unresolved = ValidateCallbacks(brokenXml, typeof(GanttRibbon));

        Assert.Contains("DoesNotExist", unresolved);
    }

    /// <summary>
    /// Every toggleButton onAction resolves to a method that declares the pressed argument.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Live Excel reported "Callback signature mismatch: OnToggleCellLockClick" the first time the
    /// 2026-11-09 toggle buttons were pressed. RibbonX gives a <c>toggleButton</c> a different
    /// <c>onAction</c> signature from a <c>button</c>: <c>void (IRibbonControl, bool isPressed)</c>
    /// versus <c>void (IRibbonControl)</c>. <see cref="ValidateCallbacks"/> only resolves NAMES, so the
    /// name-only contract above passed while every click failed.
    /// </para>
    /// <para>
    /// This is the pin for that class of defect: the element type decides the arity, so the assertion
    /// is made per element type rather than per callback name. The positive control below keeps the
    /// loop from passing vacuously if the toggles are ever removed.
    /// </para>
    /// </remarks>
    [Fact]
    public void Toggle_button_callbacks_declare_the_pressed_argument()
    {
        string xml = Ribbon.GetCustomUI(WorkbookRibbonId)!;
        XDocument doc = XDocument.Parse(xml);
        XNamespace ns = NamespaceCustomUI2010;

        List<XElement> toggles = [.. doc.Descendants(ns + "toggleButton")];
        Assert.NotEmpty(toggles);

        List<string> mismatched = [];
        foreach (XElement toggle in toggles)
        {
            string? callback = toggle.Attribute("onAction")?.Value;
            if (string.IsNullOrEmpty(callback))
            {
                continue;
            }

            System.Reflection.MethodInfo? method = typeof(GanttRibbon).GetMethod(callback);
            System.Reflection.ParameterInfo[] parameters =
                method?.GetParameters() ?? [];

            bool declaresPressed = parameters.Length == 2
                && parameters[0].ParameterType == typeof(IRibbonControl)
                && parameters[1].ParameterType == typeof(bool);
            if (!declaresPressed)
            {
                mismatched.Add(callback);
            }
        }

        Assert.Empty(mismatched);
    }

    [Fact]
    public void Ribbon_declares_getEnabled_on_exactly_the_gated_controls()
    {
        // R2.2 adds the Initialise sheet button to the Data group and R2.6 adds
        // Validate beside it; both are gated on the same workbook fact as
        // Diagnostics. The set is pinned here so adding or dropping a gated
        // control fails the contract instead of silently changing the surface.
        // R5.1 adds the plot start/finish date boxes (gated on workbook +
        // initialised + not AUTO) and R5.2 adds the grpPlotTimeScale and
        // grpPlotSize groups: the scale dropdown, the month label-form
        // dropdown beside it, the margin combobox and its companion
        // centimetre box, and the four preset buttons are all gated on
        // workbook + initialised, so they join the pinned set. The two
        // read-only width/height displays are permanently disabled through
        // the dedicated GetPlotDimensionsEnabled getter (work item R5.2 D4),
        // so they carry NO getEnabled on the shared truth table and are
        // pinned separately below. The query groups by element type (buttons,
        // then dropdowns, then edit boxes), each in document order.
        string xml = Ribbon.GetCustomUI(WorkbookRibbonId)!;
        XDocument doc = XDocument.Parse(xml);
        XNamespace ns = NamespaceCustomUI2010;

        // toggleButton is queried too: the 2026-11-09 CHART and CELL lock toggles are toggle buttons
        // and declare getEnabled on the same workbook+initialised truth table. Leaving them out of the
        // query would let a gated toggle be added or dropped without this pin noticing. They therefore
        // appear in the pinned list after every <button>, which is why they are listed there rather
        // than beside the buttons they sit next to in the XML.
        // R5.18h's column control became a `menu` (owner ruling 2026-11-14), so menus
        // join the query between dropdowns and edit boxes; a gated menu dropped or added
        // without this pin noticing is exactly what the pin exists to stop.
        var gatedIds = doc.Descendants(ns + "button")
            .Concat(doc.Descendants(ns + "toggleButton"))
            .Concat(doc.Descendants(ns + "dropDown"))
            .Concat(doc.Descendants(ns + "menu"))
            .Concat(doc.Descendants(ns + "editBox"))
            .Where(button => button.Attribute("getEnabled") is not null
                && button.Attribute("getEnabled")?.Value != "GetPlotDimensionsEnabled")
            .Select(button => button.Attribute("id")?.Value)
            .ToList();

        Assert.Equal(
            [
                RibbonControlIds.InitialiseSheet,
                // R5.16: Validate now sits inside the Fix menu, after the
                // primary Repair button in document order.
                RibbonControlIds.RepairConfig,
                RibbonControlIds.ValidateSheet,
                RibbonControlIds.AddActivity,
                // R5.16: the Critical Interval item is the only entity menu item
                // with its own gate, so it appears next in document order.
                RibbonControlIds.CriticalIntervalMenuItem,
                RibbonControlIds.AddMilestone,
                RibbonControlIds.AddDelineator,
                RibbonControlIds.Remove,
                // R5.16a: the merged Outline group. Promote/Demote then Up/Down
                // then Expand/Collapse, in document order.
                RibbonControlIds.PromoteToParent,
                RibbonControlIds.DemoteToChild,
                RibbonControlIds.MoveUp,
                RibbonControlIds.MoveDown,
                RibbonControlIds.ExpandAllGroups,
                RibbonControlIds.CollapseAllGroups,
                RibbonControlIds.ThemeSettings,
                RibbonControlIds.ChartTitle,
                // R5.1a: each calendar button carries its edit box's exact gate,
                // so an enabled picker can never sit beside a disabled box.
                RibbonControlIds.PlotStartCalendar,
                RibbonControlIds.PlotFinishCalendar,
                RibbonControlIds.RefreshSheet,
                RibbonControlIds.Warnings,
                // R6.8: Copy Editable sits in the new Export group, between the Sheet and
                // Diagnostics groups in document order, and carries the same R5.16b gate.
                RibbonControlIds.CopyEditable,
                // R7.6: Send to PowerPoint follows Copy Editable in the Export group's
                // document order and shares its corrected Mode == Ready gate.
                RibbonControlIds.SendToPowerPoint,
                // R8.8: Export PNG is the Export group's split button's main action, gated on
                // the same Ready + active-Gantt-sheet fact as Copy Editable; the command
                // ensures chart freshness itself.
                RibbonControlIds.ExportPng,
                RibbonControlIds.Diagnostics,
                RibbonControlIds.OpenLog,
                // Owner ruling 2026-11-09: the CHART and CELL lock toggles replace the shape pair.
                // They are gated on workbook + initialised alone — a toggle must always be reachable,
                // because its state is what the user came to change. R5.18e stacks them with the
                // Warnings button in the Sheet group's vertical box, cells before chart, so the
                // document order follows the owner's requested stack.
                // R5.20 adds the automatic-refresh toggle inside the Refresh split button's menu.
                // It is a toggleButton like the lock pair, so the query picks it up, and it is
                // gated on the same facts for the same reason: an off switch must stay pressable
                // or the user could never turn it back on. It appears first because the Sheet
                // group precedes the Diagnostics group in document order.
                RibbonControlIds.ToggleAutoRefresh,
                RibbonControlIds.ToggleCellLock,
                RibbonControlIds.ToggleChartLock,
                // R8.8: the two PNG-export target toggles live in the Export PNG split
                // button's menu, after the Sheet group in document order. They are gated on
                // the same facts so the target can always be switched.
                RibbonControlIds.ExportPngTargetClipboard,
                RibbonControlIds.ExportPngTargetFile,
                // R5.5: the description-label dropdown is the first
                // selection-gated control — it shares the workbook+initialised
                // part of the truth table and adds the single-row condition
                // inside it. R5.14 adds the parallel date-label dropdown
                // beside it in the Labels group, gated on the same facts.
                // R5.6a adds the custom-style dropdown after it (document
                // order), gated on a single style-key row AND at least one
                // user-created style key existing.
                RibbonControlIds.DescriptionLabel,
                RibbonControlIds.DateLabel,
                RibbonControlIds.CustomStyle,
                // R5.16a: the merged Chart group holds the dates and the scale;
                // the Size group holds the size dropdown and the margin pair.
                // Dropdowns are listed before edit boxes, and menus between the
                // two, because the query groups by element type rather than by
                // document order.
                RibbonControlIds.PlotTimeScale,
                RibbonControlIds.SizeDropdown,
                RibbonControlIds.Margin,
                RibbonControlIds.PlotStartDate,
                RibbonControlIds.PlotFinishDate,
                // R8.9: the custom width box is gated on the same workbook +
                // initialised fact plus Custom mode itself, so it joins the pinned
                // set in document order before the margin box. The height box was
                // withdrawn with the width-only ruling.
                RibbonControlIds.SizeWidthCm,
                RibbonControlIds.MarginCm,
            ],
            gatedIds);
    }

    /// <summary>
    /// No control or item id appears twice in the tab.
    /// </summary>
    /// <remarks>
    /// RibbonX ids are **global across the custom UI**, not scoped to their container, so
    /// two dropdowns may not both offer an item with the same id. The 2026-11-27 live
    /// load failed with "Custom UI Runtime Error ... The ID 'Custom' is duplicated" for
    /// exactly that reason: the Size dropdown's Custom item reused the margin dropdown's
    /// existing id. Every other ribbon pin checks an expected id, so only a global
    /// uniqueness check catches a collision between two lists.
    /// </remarks>
    [Fact]
    public void Ribbon_ids_are_unique_across_the_whole_tab()
    {
        XDocument doc = XDocument.Parse(Ribbon.GetCustomUI(WorkbookRibbonId)!);
        XNamespace ns = NamespaceCustomUI2010;

        string[] duplicates =
        [
            .. doc.Descendants()
                .Where(element => element.Name.Namespace == ns)
                .Select(element => (string?)element.Attribute("id"))
                .Where(id => !string.IsNullOrEmpty(id))
                .GroupBy(id => id!, StringComparer.Ordinal)
                .Where(group => group.Count() > 1)
                .Select(group => group.Key),
        ];

        Assert.Empty(duplicates);
    }

    [Fact]
    public void Plot_dimension_displays_are_read_only_labels()
    {
        // R5.16a: the width/height readouts are labelControls, not disabled
        // edit boxes. A label has no input affordance at all, which is what
        // "read-only output" actually means on a ribbon; the old disabled
        // boxes read as broken fields.
        string xml = Ribbon.GetCustomUI(WorkbookRibbonId)!;
        XDocument doc = XDocument.Parse(xml);
        XNamespace ns = NamespaceCustomUI2010;

        Assert.Equal(
            ["GetPlotWidthLabel", "GetPlotHeightLabel"],
            doc.Descendants(ns + "labelControl")
                .Select(element => element.Attribute("getLabel")?.Value)
                .Where(value => value is "GetPlotWidthLabel" or "GetPlotHeightLabel"));
        Assert.DoesNotContain(
            doc.Descendants(ns + "editBox"),
            element => element.Attribute("id")?.Value is "edtPlotWidth" or "edtPlotHeight");

        var service = new RibbonStateService();
        Assert.Equal("W", GanttRibbon.GetPlotWidthLabel(null, service));
        Assert.Equal("H", GanttRibbon.GetPlotHeightLabel(null, service));
    }

    [Fact]
    public void The_columns_control_is_a_split_button_whose_menu_rows_are_the_columns()
    {
        // R5.22/ADR-0065 D3: ONE split button. The primary toggles all six; the menu keeps
        // the per-column toggles. Office 2010's CT_SplitButton puts the primary `button`
        // first and the `menu` second, and the parent carries `size`/`getEnabled`.
        string xml = Ribbon.GetCustomUI(WorkbookRibbonId)!;
        XDocument doc = XDocument.Parse(xml);
        XNamespace ns = NamespaceCustomUI2010;

        XElement split = Assert.Single(
            doc.Descendants(ns + "splitButton"),
            element => element.Attribute("id")?.Value == RibbonControlIds.ColumnsSplitButton);

        Assert.Equal("large", split.Attribute("size")?.Value);
        Assert.NotNull(split.Attribute("getEnabled"));

        // The child order is part of the schema: button (or toggleButton) then menu.
        Assert.Equal(
            ["button", "menu"],
            split.Elements().Select(element => element.Name.LocalName));

        XElement primary = Assert.Single(split.Elements(ns + "button"));
        Assert.Equal(RibbonControlIds.ColumnVisibilityAll, primary.Attribute("id")?.Value);
        Assert.Equal("GetColumnVisibilityAllLabel", primary.Attribute("getLabel")?.Value);
        Assert.Equal("GetGlyph", primary.Attribute("getImage")?.Value);
        Assert.Equal("OnColumnVisibilityToggleAll", primary.Attribute("onAction")?.Value);

        // The menu is the drop-down, not a ribbon control: no size and no second gate.
        XElement menu = Assert.Single(split.Elements(ns + "menu"));
        Assert.Equal(RibbonControlIds.ColumnsMenu, menu.Attribute("id")?.Value);
        Assert.Null(menu.Attribute("size"));
        Assert.Null(menu.Attribute("getEnabled"));

        string[] expected =
        [
            .. GanttColumnVisibility.ToggleableColumnNames
                .Select(name => RibbonControlIds.ColumnItemPrefix + GanttColumnVisibility.RibbonItemId(name)),
        ];
        Assert.Equal(
            expected,
            menu.Elements(ns + "button").Select(element => element.Attribute("id")!.Value));

        Assert.All(
            menu.Elements(ns + "button"),
            row =>
            {
                Assert.Equal("GetColumnVisibilityLabel", row.Attribute("getLabel")?.Value);
                Assert.Equal("GetColumnVisibilityImage", row.Attribute("getImage")?.Value);
                Assert.Equal("OnColumnVisibilityToggle", row.Attribute("onAction")?.Value);
            });

        // The old combo is gone: a leftover dropDown would render beside the split button.
        Assert.DoesNotContain(
            doc.Descendants(ns + "dropDown"),
            element => element.Attribute("id")?.Value == "ddnColumnVisibility");
    }

    [Fact]
    public void The_columns_split_button_carries_the_single_gate()
    {
        // A splitButton is the ribbon control now, so its getEnabled is the one gate; the
        // primary and the menu rows inherit it. The broader gated-control pin queries only
        // buttons/dropdowns/editBoxes/menus, so this fact owns the splitButton placement.
        string xml = Ribbon.GetCustomUI(WorkbookRibbonId)!;
        XDocument doc = XDocument.Parse(xml);
        XNamespace ns = NamespaceCustomUI2010;

        XElement split = Assert.Single(
            doc.Descendants(ns + "splitButton"),
            element => element.Attribute("id")?.Value == RibbonControlIds.ColumnsSplitButton);

        _ = Assert.Single(split.Attributes("getEnabled"));
        Assert.Null(split.Elements(ns + "menu").Single().Attribute("getEnabled"));
    }

    [Fact]
    public void Ribbon_declares_getEnabled_on_each_control_exactly_once()
    {
        // Each gated control declares getEnabled exactly once: a duplicated
        // attribute on one control would make the count above pass while the
        // ribbon rendered with a duplicate callback.
        string xml = Ribbon.GetCustomUI(WorkbookRibbonId)!;
        XDocument doc = XDocument.Parse(xml);
        XNamespace ns = NamespaceCustomUI2010;

        foreach (string id in new[]
        {
            RibbonControlIds.ColumnsSplitButton,
            RibbonControlIds.Diagnostics,
            RibbonControlIds.OpenLog,
            RibbonControlIds.InitialiseSheet,
            RibbonControlIds.ValidateSheet,
            RibbonControlIds.RepairConfig,
            RibbonControlIds.AddActivity,
            RibbonControlIds.AddMilestone,
            RibbonControlIds.AddDelineator,
            RibbonControlIds.PromoteToParent,
            RibbonControlIds.DemoteToChild,
            RibbonControlIds.MoveUp,
            RibbonControlIds.MoveDown,
            RibbonControlIds.RefreshSheet,
            RibbonControlIds.ThemeSettings,
            RibbonControlIds.ChartTitle,
            RibbonControlIds.PlotStartDate,
            RibbonControlIds.PlotFinishDate,
            RibbonControlIds.PlotTimeScale,
            RibbonControlIds.Margin,
            RibbonControlIds.MarginCm,
            RibbonControlIds.SizeDropdown,
        })
        {
            var matches = doc.Descendants(ns + "button")
                .Concat(doc.Descendants(ns + "splitButton"))
                .Concat(doc.Descendants(ns + "dropDown"))
                .Concat(doc.Descendants(ns + "editBox"))
                .Where(b => b.Attribute("id")?.Value == id)
                .Select(b => b.Attribute("getEnabled"))
                .Count(attr => attr is not null);

            Assert.Equal(1, matches);
        }
    }

    [Fact]
    public void Ribbon_buttons_are_served_by_the_themed_glyph_callback()
    {
        // imageMso fails silently on unknown IDs (no error, no log): the
        // Diagnostics button showed no icon with imageMso="Information",
        // which is absent from the Office 2010 Icons Gallery workbook
        // (customUI14.xml: 7344 unique IDs, no "Information"). Pin
        // evidence-backed IDs so a typo regresses to a failing test
        // instead of a silently missing icon.
        //
        // Two bases, both required in practice: the ID must be present in
        // the Office 2010 Icons Gallery (Office2010IconsGallery.docx, whose
        // embedded customUI14.xml lists 7344 unique imageMso IDs) and name a
        // command in the Excel idMso table of [MS-CUSTOMUI]-250218
        // (`.microsoft/` reference, git-ignored).
        // imageMso="OfficeDiagnostics" was a valid gallery ID but absent
        // from the Excel table and rendered nothing in F5. FileOpen and Help
        // satisfy both and render (Help glyph confirmed by a human operator
        // in F5, 2026-09-16).
        //
        // TableInsertExcel satisfies both: present in the gallery, and the
        // Excel idMso row idMso=TableInsertExcel, control=button,
        // label="Table" — the Excel Insert-Table command, the semantically
        // exact match for "Initialise sheet", which creates tblGanttData.
        // Note the id is TableInsertExcel, not its UI label "Table", which
        // is itself absent from the gallery. The gallery check is validated
        // against the four documented outcomes that file reproduces
        // (Information absent, OfficeDiagnostics present, Help present,
        // FileOpen present). Evidence and the remaining F5 glyph step:
        // docs/work-items/R2.2-initialise-sheet.md decision D8.
        //
        // The four preset buttons use the orientation icons (the user's
        // mapping: A4 portrait/4:3 -> portrait glyph, A4 landscape/16:9 ->
        // landscape glyph). Both IDs satisfy the two bases: the gallery
        // lists them as button entries (customUI14.xml: imageMso=
        // PageOrientationPortrait / PageOrientationLandscape), and the
        // Excel idMso table rows are idMso=PageOrientationPortrait,
        // control=toggleButton, label="Portrait" and idMso=
        // PageOrientationLandscape, control=toggleButton, label=
        // "Landscape". The icon is what imageMso borrows, so a
        // toggleButton id is valid on a button element — the gallery
        // itself renders these on <button>. The previous values failed
        // this test's protocol: "Presentation" is absent from the gallery
        // entirely (0 occurrences; Excel reported "Unknown Office control
        // ID: Presentation" at load), and "PageSetup" is gallery-present
        // but absent from the Excel idMso table — the OfficeDiagnostics
        // failure mode. The orientation IDs replace both, satisfy the
        // two bases, and are orientation-specific.
        string xml = Ribbon.GetCustomUI(WorkbookRibbonId)!;
        XDocument doc = XDocument.Parse(xml);
        XNamespace ns = NamespaceCustomUI2010;

        string? ImageCallback(string id) =>
            doc.Descendants(ns + "button")
                .First(b => b.Attribute("id")?.Value == id)
                .Attribute("getImage")?.Value;

        // R5.19: every control-level action names the one themed callback. What each id
        // then resolves — a Lucide slug at a declared size, present in both palettes — is
        // pinned once in RibbonGlyphAssetsTests against the generated catalogue, so this
        // test owns the wiring rather than a second copy of the map.
        foreach (string id in new[]
                 {
                     RibbonControlIds.Diagnostics,
                     RibbonControlIds.OpenLog,
                     RibbonControlIds.InitialiseSheet,
                     RibbonControlIds.RepairConfig,
                     RibbonControlIds.ValidateSheet,
                     RibbonControlIds.Remove,
                     RibbonControlIds.PromoteToParent,
                     RibbonControlIds.DemoteToChild,
                     RibbonControlIds.MoveUp,
                     RibbonControlIds.MoveDown,
                     RibbonControlIds.ExpandAllGroups,
                     RibbonControlIds.CollapseAllGroups,
                     RibbonControlIds.ThemeSettings,
                     RibbonControlIds.ChartTitle,
                     RibbonControlIds.RefreshSheet,
                     RibbonControlIds.Warnings,
                 })
        {
            Assert.Equal("GetGlyph", ImageCallback(id));
        }

        // The three add primaries show a glyph that follows the label beside it, so they
        // name their own callback; RibbonGlyphAssetsTests pins that pairing and the slugs.
        foreach (string id in new[]
                 {
                     RibbonControlIds.AddActivity,
                     RibbonControlIds.AddMilestone,
                     RibbonControlIds.AddDelineator,
                 })
        {
            Assert.Equal("GetAddEntityPrimaryImage", ImageCallback(id));
        }

        // The migration is complete only when no control still falls back to a gallery id.
        Assert.DoesNotContain(
            doc.Descendants(ns + "button"),
            button => button.Attribute("imageMso") is not null);
    }

    [Fact]
    public void Control_level_glyphs_are_unique_except_the_plot_date_calendar_pair()
    {
        // R5.16a D4 still holds: two different actions must not look the same. Menu rows
        // are exempt, because a dropdown legitimately repeats a family glyph and the row
        // label carries the variant. The comparison runs on the generated catalogue, so it
        // measures the glyph actually served rather than the attribute in the XML.
        //
        // R5.1a owner ruling: the two plot-date calendar buttons deliberately share one
        // glyph. The label column and the supertip name the end, and sharing the asset is
        // what keeps one decode for both. It is the ONLY allowed duplicate, pinned by the
        // single-duplicate assertion below rather than by dropping the rule.
        string xml = Ribbon.GetCustomUI(WorkbookRibbonId)!;
        XDocument doc = XDocument.Parse(xml);
        XNamespace ns = NamespaceCustomUI2010;

        (string Id, string Glyph)[] controls =
        [
            .. doc.Descendants(ns + "button")
                .Where(button => !button.Ancestors(ns + "menu").Any())
                .Select(button => (
                    Id: button.Attribute("id")?.Value,
                    Glyph: button.Attribute("id")?.Value is string id
                        && RibbonGlyphCatalog.TryGet(id, out RibbonGlyphCatalog.Entry entry)
                            ? entry.Glyph
                            : null))
                .Where(pair => pair.Id is not null && pair.Glyph is not null)
                .Select(pair => (pair.Id!, pair.Glyph!)),
        ];

        Assert.NotEmpty(controls);

        string[] expectedShared =
        [
            .. new[] { RibbonControlIds.PlotFinishCalendar, RibbonControlIds.PlotStartCalendar }
                .OrderBy(id => id, StringComparer.Ordinal),
        ];

        List<string[]> duplicates =
        [
            .. controls
                .GroupBy(pair => pair.Glyph, StringComparer.Ordinal)
                .Where(group => group.Count() > 1)
                .Select(group => group.Select(pair => pair.Id).OrderBy(id => id, StringComparer.Ordinal).ToArray()),
        ];

        Assert.Single(duplicates);
        Assert.Equal(expectedShared, duplicates[0]);
    }

    [Fact]
    public void Every_actionable_control_carries_a_unique_keytip_and_accessible_description()
    {
        // R5.9 D1/D2: every actionable control is reachable by keyboard and
        // announced by a screen reader. Excel ignores an unknown or duplicate
        // keytip silently, so presence and uniqueness are pinned here rather
        // than left to a live review.
        string xml = Ribbon.GetCustomUI(WorkbookRibbonId)!;
        XDocument doc = XDocument.Parse(xml);

        List<XElement> actionable = [.. doc.Descendants()
            .Where(element => element.Name.LocalName is "button" or "dropDown" or "checkBox" or "editBox")];
        Assert.NotEmpty(actionable);

        foreach (XElement control in actionable)
        {
            string id = control.Attribute("id")?.Value ?? "(missing id)";
            Assert.False(
                string.IsNullOrWhiteSpace(control.Attribute("keytip")?.Value),
                $"{id} declares no keytip.");
            Assert.False(
                string.IsNullOrWhiteSpace(control.Attribute("screentip")?.Value),
                $"{id} declares no screentip.");
            Assert.False(
                string.IsNullOrWhiteSpace(control.Attribute("supertip")?.Value),
                $"{id} declares no supertip.");
            Assert.InRange(control.Attribute("keytip")!.Value.Length, 1, 3);
        }

        string[] keytips = [.. actionable.Select(control => control.Attribute("keytip")!.Value)];
        Assert.Equal(keytips.Length, keytips.Distinct(StringComparer.OrdinalIgnoreCase).Count());
    }

    [Fact]
    public void Every_button_and_unlabelled_dropdown_carries_a_visible_label()
    {
        // The visible label is the control's accessible name; the screentip and
        // supertip are its description. A button without a label is announced
        // as an unnamed button, and a dropdown without a companion
        // labelControl would have no visible name at all.
        string xml = Ribbon.GetCustomUI(WorkbookRibbonId)!;
        XDocument doc = XDocument.Parse(xml);
        XNamespace ns = NamespaceCustomUI2010;

        // R5.16: a menu item's visible label may be projected through getLabel
        // from the catalogue, so a static label is no longer the only source.
        Assert.Empty(ButtonsMissingVisibleLabel(xml));

        foreach (XElement dropdown in doc.Descendants(ns + "dropDown"))
        {
            string? id = dropdown.Attribute("id")?.Value;
            string? companionId = id is null ? null : "lbl" + id["ddn".Length..];
            bool hasCompanionLabel = companionId is not null
                && doc.Descendants(ns + "labelControl")
                    .Any(label => label.Attribute("id")?.Value == companionId);
            Assert.True(
                !string.IsNullOrWhiteSpace(dropdown.Attribute("label")?.Value) || hasCompanionLabel,
                $"Dropdown {id} has neither a label nor a companion labelControl.");
        }
    }
    [Fact]
    public void The_visible_label_validator_flags_a_button_with_neither_label_nor_getLabel()
    {
        const string xml = """
            <customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
              <ribbon><tabs><tab id="tab"><group id="grp">
                <button id="btnLabelled" label="Labelled" keytip="L"/>
                <button id="btnProjected" getLabel="GetEntityMenuLabel" keytip="P"/>
                <button id="btnBare" keytip="B"/>
              </group></tab></tabs></ribbon>
            </customUI>
            """;

        Assert.Equal(["btnBare"], ButtonsMissingVisibleLabel(xml));
    }

    [Fact]
    public void Entity_menu_labels_are_projected_from_the_catalogue()
    {
        Assert.Equal("As-Planned Activity", GanttRibbon.GetEntityMenuLabel("entAsPlannedActivity"));
        Assert.Equal("Critical Interval", GanttRibbon.GetEntityMenuLabel("entCriticalInterval"));
        Assert.Equal("Delineator", GanttRibbon.GetEntityMenuLabel("entDelineator"));

        // Unknown, non-prefixed, and absent ids project to empty, never a throw.
        Assert.Equal(string.Empty, GanttRibbon.GetEntityMenuLabel("entNotAType"));
        Assert.Equal(string.Empty, GanttRibbon.GetEntityMenuLabel(RibbonControlIds.Remove));
        Assert.Equal(string.Empty, GanttRibbon.GetEntityMenuLabel(null));
    }

    [Fact]
    public void Entity_menu_ids_require_an_exact_catalogue_name()
    {
        // A numeric or combined enum value must not smuggle in a type.
        Assert.False(GanttRibbon.TryParseEntityMenuId("ent1", out _));
        Assert.False(GanttRibbon.TryParseEntityMenuId("ent", out _));
        Assert.False(GanttRibbon.TryParseEntityMenuId("entAsPlannedActivityX", out _));
        Assert.False(GanttRibbon.TryParseEntityMenuId(null, out _));

        Assert.True(GanttRibbon.TryParseEntityMenuId("entAsPlannedActivity", out GanttEntityType type));
        Assert.Equal(GanttEntityType.AsPlannedActivity, type);
    }

    [Fact]
    public void The_three_entity_menus_partition_the_catalogue()
    {
        XDocument doc = XDocument.Parse(Ribbon.GetCustomUI(WorkbookRibbonId)!);
        XNamespace ns = NamespaceCustomUI2010;

        List<GanttEntityType> Types(string menuId) =>
        [
            .. doc.Descendants(ns + "menu")
                .Single(menu => (string?)menu.Attribute("id") == menuId)
                .Elements(ns + "button")
                .Select(button =>
                {
                    Assert.True(GanttRibbon.TryParseEntityMenuId(
                        (string?)button.Attribute("id"), out GanttEntityType type));
                    return type;
                }),
        ];

        List<GanttEntityType> activities = Types(RibbonControlIds.AddActivityMenu);
        List<GanttEntityType> milestones = Types(RibbonControlIds.AddMilestoneMenu);
        List<GanttEntityType> elements = Types(RibbonControlIds.AddElementsMenu);

        Assert.Equal(
            Enum.GetValues<GanttEntityType>()
                .Where(type => EntityTypeCatalog.GetDefinition(type)?.Kind == EntityKind.Span)
                .Order(),
            activities.Order());
        Assert.Equal(
            Enum.GetValues<GanttEntityType>()
                .Where(type => EntityTypeCatalog.GetDefinition(type)?.Kind == EntityKind.Milestone)
                .Order(),
            milestones.Order());
        Assert.Equal(
            new[] { GanttEntityType.Delineator, GanttEntityType.Spacer, GanttEntityType.Splitter }.Order(),
            elements.Order());

        List<GanttEntityType> union = [.. activities, .. milestones, .. elements];
        Assert.Equal(EntityTypeCatalog.TypeCount, union.Distinct().Count());
        Assert.Equal(EntityTypeCatalog.TypeCount, union.Count);
    }

    [Fact]
    public void Ribbon_help_text_is_offline_only()
    {
        // R5.9 D3 / the offline invariant: no network link may appear in the
        // ribbon help text. A supertip is local text; a URL would be a hidden
        // network dependency. The XML namespace itself is excluded by reading
        // attribute values only.
        string xml = Ribbon.GetCustomUI(WorkbookRibbonId)!;
        XDocument doc = XDocument.Parse(xml);

        foreach (XElement element in doc.Descendants())
        {
            foreach (string attributeName in new[] { "keytip", "screentip", "supertip", "label" })
            {
                string? value = element.Attribute(attributeName)?.Value;
                if (value is null)
                {
                    continue;
                }

                Assert.DoesNotContain("http://", value, StringComparison.OrdinalIgnoreCase);
                Assert.DoesNotContain("https://", value, StringComparison.OrdinalIgnoreCase);
            }
        }
    }

    /// <summary>
    /// The Plot Area group owns no Apply control: the Apply-button
    /// contract is replaced by per-end persistence, and the ribbon state
    /// does not consult an obsolete control ID.
    /// </summary>
    [Fact]
    public void PlotArea_group_has_no_apply_button()
    {
        string xml = Ribbon.GetCustomUI(WorkbookRibbonId)!;
        XDocument doc = XDocument.Parse(xml);
        XNamespace ns = NamespaceCustomUI2010;

        var buttonIds = doc.Descendants(ns + "button")
            .Select(button => button.Attribute("id")?.Value)
            .ToList();

        Assert.DoesNotContain("btnApplyPlotSettings", buttonIds, StringComparer.Ordinal);
        Assert.DoesNotContain("btnPlotSettingsAuto", buttonIds, StringComparer.Ordinal);
        Assert.DoesNotContain("btnPlotSettingsExplicit", buttonIds, StringComparer.Ordinal);
    }

    // ---------------------------------------------------------------------
    // R8.8 — the Export group's action surface.
    // ---------------------------------------------------------------------

    /// <summary>
    /// R8.8 (owner ruling 2026-11-23): the Export group carries the three actions and no
    /// width input. The width/unit/preset/preview surface is withdrawn, so an edit box,
    /// dropdown or preview label reappearing here fails this pin.
    /// </summary>
    [Fact]
    public void The_export_group_carries_the_actions_and_no_width_input()
    {
        XDocument doc = XDocument.Parse(Ribbon.GetCustomUI(WorkbookRibbonId)!);
        XNamespace ns = NamespaceCustomUI2010;
        XElement group = doc.Descendants(ns + "group")
            .Single(element => (string?)element.Attribute("id") == RibbonControlIds.ExportGroup);

        Assert.Empty(group.Descendants(ns + "editBox"));
        Assert.Empty(group.Descendants(ns + "dropDown"));
        Assert.Empty(group.Descendants(ns + "labelControl"));

        XElement export = group.Descendants(ns + "button")
            .Single(element => (string?)element.Attribute("id") == RibbonControlIds.ExportPng);
        Assert.Equal("OnExportPngClick", (string?)export.Attribute("onAction"));
        Assert.Equal("GetEnabled", (string?)export.Attribute("getEnabled"));
    }

    /// <summary>
    /// R8.8: the Export PNG action is a split button whose menu carries the two radio-style
    /// target toggles, each wired to its pressed getter and the shared onAction, and whose
    /// main button's label states the current target.
    /// </summary>
    [Fact]
    public void The_export_png_action_is_a_split_button_with_a_target_menu()
    {
        XDocument doc = XDocument.Parse(Ribbon.GetCustomUI(WorkbookRibbonId)!);
        XNamespace ns = NamespaceCustomUI2010;
        XElement split = doc.Descendants(ns + "splitButton")
            .Single(element => (string?)element.Attribute("id") == RibbonControlIds.ExportPngSplit);

        XElement primary = split.Descendants(ns + "button")
            .Single(element => (string?)element.Attribute("id") == RibbonControlIds.ExportPng);
        Assert.Equal("GetExportPngLabel", (string?)primary.Attribute("getLabel"));
        Assert.Equal("OnExportPngClick", (string?)primary.Attribute("onAction"));

        XElement clipboard = split.Descendants(ns + "toggleButton")
            .Single(element => (string?)element.Attribute("id") == RibbonControlIds.ExportPngTargetClipboard);
        Assert.Equal("GetExportPngClipboardPressed", (string?)clipboard.Attribute("getPressed"));
        Assert.Equal("OnExportPngTargetClick", (string?)clipboard.Attribute("onAction"));

        XElement file = split.Descendants(ns + "toggleButton")
            .Single(element => (string?)element.Attribute("id") == RibbonControlIds.ExportPngTargetFile);
        Assert.Equal("GetExportPngFilePressed", (string?)file.Attribute("getPressed"));
        Assert.Equal("OnExportPngTargetClick", (string?)file.Attribute("onAction"));
    }
}
