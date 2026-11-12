namespace GanttCreator.AddIn;

/// <summary>
/// The stable Ribbon control IDs for the controls whose enabled state is driven
/// by <see cref="RibbonState"/>.
/// </summary>
/// <remarks>
/// These are the single source for the IDs, shared by the state getter truth
/// table and the RibbonX contract tests (which assert the shipped
/// <c>Ribbon.xml</c> declares <c>getEnabled</c> on exactly these controls). A
/// future dynamic control is added both here and to the RibbonX document, or the
/// XML test fails.
/// </remarks>
internal static class RibbonControlIds
{
    /// <summary>The Diagnostics button (<c>getEnabled</c> driven by workbook presence).</summary>
    internal const string Diagnostics = "btnDiagnostics";

    /// <summary>The Open log file button (<c>getEnabled</c> driven by log availability).</summary>
    internal const string OpenLog = "btnOpenLog";

    /// <summary>The Initialise sheet button (<c>getEnabled</c> driven by workbook presence).</summary>
    internal const string InitialiseSheet = "btnInitialiseSheet";

    /// <summary>The Validate button (<c>getEnabled</c> driven by workbook presence).</summary>
    internal const string ValidateSheet = "btnValidateSheet";

    /// <summary>
    /// The Refresh chart button (<c>getEnabled</c> driven by workbook presence).
    /// </summary>
    /// <remarks>
    /// Gated on workbook presence alone, exactly as Validate is. A stricter gate —
    /// "an initialised table exists" — was considered and rejected: it would grey the
    /// button on a workbook the user is halfway through setting up, and the command
    /// already refuses with an actionable message when the table is missing. A
    /// button that is enabled and explains itself beats a button that silently
    /// disables itself for a condition the user cannot see.
    /// </remarks>
    internal const string RefreshSheet = "btnRefreshSheet";

    /// <summary>The explicit scene-warning list button.</summary>
    internal const string Warnings = "btnWarnings";

    /// <summary>The Repair configuration button (<c>getEnabled</c> driven by workbook presence).</summary>
    internal const string RepairConfig = "btnRepairConfig";

    /// <summary>The Add activity button (<c>getEnabled</c> driven by workbook presence).</summary>
    internal const string AddActivity = "btnAddActivity";

    /// <summary>The Add milestone button (<c>getEnabled</c> driven by workbook presence).</summary>
    internal const string AddMilestone = "btnAddMilestone";

    /// <summary>The Add delineator button (<c>getEnabled</c> driven by workbook presence).</summary>
    internal const string AddDelineator = "btnAddDelineator";

    // R5.16: the four Data-group split buttons and their menus. The primary
    // button of each pair keeps its own id above, so the state truth table and
    // the pinned gated-control contract keep naming the control the user clicks.

    /// <summary>The Fix split button (primary Repair configuration, menu Validate).</summary>
    internal const string FixSplit = "spbFix";

    /// <summary>The Fix menu.</summary>
    internal const string FixMenu = "mnuFix";

    /// <summary>The Add activity split button.</summary>
    internal const string AddActivitySplit = "spbAddActivity";

    /// <summary>The Add activity menu.</summary>
    internal const string AddActivityMenu = "mnuAddActivity";

    /// <summary>The Add milestone split button.</summary>
    internal const string AddMilestoneSplit = "spbAddMilestone";

    /// <summary>The Add milestone menu.</summary>
    internal const string AddMilestoneMenu = "mnuAddMilestone";

    /// <summary>The Additional elements split button.</summary>
    internal const string AddElementsSplit = "spbAddElements";

    /// <summary>The Additional elements menu.</summary>
    internal const string AddElementsMenu = "mnuAddElements";

    /// <summary>
    /// Prefix for every entity menu item id: <c>ent</c> + the
    /// <see cref="Core.GanttEntityType"/> name. The suffix is the durable enum
    /// name, so the id needs no second table.
    /// </summary>
    internal const string EntityMenuItemPrefix = "ent";

    /// <summary>
    /// The Critical Interval menu item (R5.16). It is the only entity item with
    /// its own gate: the table must hold a row that can own the interval.
    /// </summary>
    internal const string CriticalIntervalMenuItem = "entCriticalInterval";

    /// <summary>
    /// The CHART lock toggle (owner ruling 2026-11-09): one control whose label and pressed state
    /// reflect whether chart elements can be dragged, and whose action flips it.
    /// </summary>
    internal const string ToggleChartLock = "btnToggleChartLock";

    /// <summary>
    /// The CELL lock toggle (owner ruling 2026-11-09): one control for the sheet's cells OUTSIDE the
    /// visible table. The table itself is always editable and is not affected.
    /// </summary>
    internal const string ToggleCellLock = "btnToggleCellLock";

    /// <summary>The Promote to parent button (<c>getEnabled</c> driven by workbook + initialised).</summary>
    internal const string PromoteToParent = "btnPromoteToParent";

    /// <summary>The Demote to child button (<c>getEnabled</c> driven by workbook + initialised).</summary>
    internal const string DemoteToChild = "btnDemoteToChild";

    /// <summary>
    /// The Move up button (<c>getEnabled</c> driven by workbook + initialised).
    /// </summary>
    /// <remarks>
    /// Gated on workbook + initialised only, exactly as promote and demote are: the
    /// top of the selection's set is a quiet no-op by design (R5.4 D4), and a
    /// wrong-kind selection refuses at execution time with a message. Selection-aware
    /// gating is owed work (see the <c>TODO(R5.10)</c> note in <c>Ribbon.xml</c>).
    /// </remarks>
    internal const string MoveUp = "btnMoveUp";

    /// <summary>
    /// The Move down button (<c>getEnabled</c> driven by workbook + initialised).
    /// </summary>
    /// <remarks>
    /// The mirror of <see cref="MoveUp"/>; the bottom of the selection's set is a
    /// quiet no-op by design (R5.4 D4).
    /// </remarks>
    internal const string MoveDown = "btnMoveDown";

    /// <summary>
    /// The description-label dropdown (<c>getEnabled</c> driven by workbook
    /// + initialised + one body row selected). The ribbon's <c>Labels</c>
    /// group carries it beside a <c>Description label</c> row label; the
    /// date-label dropdown (R5.14) is its parallel.
    /// </summary>
    /// <remarks>
    /// R5.5 owner ruling 6: per-row gating ships with the label controls, so
    /// the dropdown enables only when the selection resolves to exactly one
    /// table body row (a collapsed parent qualifies and edits its own cell
    /// only — ruling 7). R5.10 reuses the same selection facts for fill/stroke.
    /// </remarks>
    internal const string DescriptionLabel = "ddnDescriptionLabel";

    /// <summary>
    /// The date-label dropdown (<c>getEnabled</c> driven by workbook
    /// + initialised + one body row selected). The ribbon's <c>Labels</c>
    /// group carries it beside a <c>Date label</c> row label; it is the
    /// parallel of <see cref="DescriptionLabel"/>.
    /// </summary>
    /// <remarks>
    /// R5.14: the item list is projected at query time from the selected
    /// row's <c>Type</c> through <c>EntityTypeCatalog.GetAllowedDateLabelPositions</c>
    /// (the capability is type-only — a Custom Activity style carries no
    /// date mode, so it cannot narrow the set). The dropdown enables only
    /// when the selection resolves to exactly one table body row, reusing
    /// the same selection facts the description-label dropdown uses.
    /// </remarks>
    internal const string DateLabel = "ddnDateLabel";

    /// <summary>
    /// The custom-style dropdown (<c>getEnabled</c> driven by workbook +
    /// initialised + a single style-key row + at least one user-created
    /// style key). R5.6a: items are the non-built-in keys of the
    /// configuration style registry, projected at query time; assigning
    /// writes the selected row's hidden <c>StyleKey</c> cell.
    /// </summary>
    internal const string CustomStyle = "ddnCustomStyle";

    /// <summary>
    /// The Expand all groups button (<c>getEnabled</c> driven by workbook +
    /// initialised).
    /// </summary>
    /// <remarks>
    /// R5.4: the Ribbon surface over the native Excel outline. It expands every
    /// parent/child group at once, which is the only thing the native
    /// per-group <c>+</c> control cannot do. Gated on workbook + initialised
    /// like every other Data-group button: a sheet with no group is a quiet
    /// no-op by design, and the command refuses with an actionable message when
    /// the table is missing.
    /// </remarks>
    internal const string ExpandAllGroups = "btnExpandAllGroups";

    /// <summary>
    /// The Collapse all groups button (<c>getEnabled</c> driven by workbook +
    /// initialised).
    /// </summary>
    /// <remarks>
    /// R5.4: the Ribbon surface over the native Excel outline. It collapses
    /// every parent/child group at once. Gated exactly like
    /// <see cref="ExpandAllGroups"/>.
    /// </remarks>
    internal const string CollapseAllGroups = "btnCollapseAllGroups";

    /// <summary>
    /// The Remove button (<c>getEnabled</c> driven by workbook +
    /// initialised). Destructive under ADR-0008: the command itself
    /// confirms before the mutator runs, so the enabled gate stays
    /// at workbook + initialised like every other Data-group button.
    /// </summary>
    internal const string Remove = "btnRemove";

    /// <summary>The Plot start AUTO checkbox (<c>getEnabled</c> driven by workbook + initialised).</summary>
    internal const string PlotStartAuto = "chkPlotStartAuto";

    /// <summary>The Plot finish AUTO checkbox (<c>getEnabled</c> driven by workbook + initialised).</summary>
    internal const string PlotFinishAuto = "chkPlotFinishAuto";

    /// <summary>The Plot start date edit box (<c>getEnabled</c> driven by workbook + initialised + not auto).</summary>
    internal const string PlotStartDate = "edtPlotStartDate";

    /// <summary>The Plot finish date edit box (<c>getEnabled</c> driven by workbook + initialised + not auto).</summary>
    internal const string PlotFinishDate = "edtPlotFinishDate";

    /// <summary>The plot time-scale dropdown (<c>getEnabled</c> driven by workbook + initialised).</summary>
    internal const string PlotTimeScale = "ddnPlotTimeScale";

    /// <summary>The plot margin combobox (<c>getEnabled</c> driven by workbook + initialised).</summary>
    internal const string Margin = "ddnMargin";

    /// <summary>The custom-margin edit box (<c>getEnabled</c> driven by the margin preset being Custom).</summary>
    internal const string MarginCm = "edtMarginCm";

    /// <summary>
    /// The plot width readout (R5.16a). A <c>labelControl</c>, never an input:
    /// the old permanently-disabled edit box read as a broken field.
    /// </summary>
    internal const string PlotWidth = "lblPlotWidth";

    /// <summary>The plot height readout (R5.16a), the parallel of <see cref="PlotWidth"/>.</summary>
    internal const string PlotHeight = "lblPlotHeight";

    /// <summary>
    /// The output-size dropdown (R5.16a). Replaces the four preset buttons and
    /// also shows the current preset; its item ids are the durable
    /// <see cref="Core.SizePreset"/> keys.
    /// </summary>
    internal const string SizeDropdown = "ddnSize";

    /// <summary>The Theme settings button (<c>getEnabled</c> driven by workbook + initialised).</summary>
    internal const string ThemeSettings = "btnThemeSettings";

    /// <summary>The Chart title button (<c>getEnabled</c> driven by workbook + initialised).</summary>
    internal const string ChartTitle = "btnChartTitle";

    // R5.16 group ids. `getVisible` is declared on every group, so the
    // visibility truth table needs a name for each one. An id that is absent
    // from RibbonState.IsVisible is HIDDEN (fail-closed): showing an unintended
    // group is the failure R5.16 removes.

    /// <summary>
    /// The add-in tab, targeted by <c>IRibbonUI.ActivateTab</c> so entering the
    /// Gantt sheet restores this tab (R5.16b, ADR-0058).
    /// </summary>
    internal const string Tab = "tabGanttCreator";

    /// <summary>The first-run group holding the large Initialise sheet button.</summary>
    internal const string GetStartedGroup = "grpGetStarted";

    /// <summary>The Data group: Repair, Add activity, Add milestone, Elements, Remove.</summary>
    internal const string DataGroup = "grpData";

    /// <summary>
    /// The Outline group (R5.16a): promote, demote, up, down, expand all,
    /// collapse all. The former Ordering group is merged into it.
    /// </summary>
    internal const string OutlineGroup = "grpOutline";

    /// <summary>The Labels group (description, date, custom style).</summary>
    internal const string LabelsGroup = "grpLabels";

    /// <summary>The Style group (theme settings, chart title).</summary>
    internal const string StyleThemeGroup = "grpStyleTheme";

    /// <summary>The Chart group: plot dates and the time scale (R5.16a merge).</summary>
    internal const string PlotAreaGroup = "grpPlotArea";

    /// <summary>The Size group: output size, margin, and the dimension readouts.</summary>
    internal const string PlotSizeGroup = "grpPlotSize";

    /// <summary>The Sheet group: refresh, warnings, and the two lock toggles.</summary>
    internal const string SheetGroup = "grpSheet";

    /// <summary>The Diagnostics group, visible in every mode.</summary>
    internal const string DiagnosticsGroup = "grpDiagnostics";
}
