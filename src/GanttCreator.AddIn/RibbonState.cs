using GanttCreator.Core;

namespace GanttCreator.AddIn;

/// <summary>
/// Which face the Ribbon shows (R5.16, ADR-0056).
/// </summary>
internal enum RibbonMode
{
    /// <summary>No add-in configuration: only Initialise and Diagnostics are offered.</summary>
    FirstRun = 0,

    /// <summary>
    /// Configuration exists but the table is missing or unreadable: the full
    /// surface is shown, Initialise is hidden, and Fix is the enabled remedy.
    /// </summary>
    NeedsRepair = 1,

    /// <summary>The table is readable: the full surface, with table-dependent controls enabled.</summary>
    Ready = 2,
}

/// <summary>
/// Immutable snapshot of the host facts the dynamic Ribbon getters read. The
/// state service captures it only from an explicit refresh, so every getter is a
/// pure read of one value — deterministic, side-effect-free, and never touching
/// Excel.
/// </summary>
/// <param name="HasActiveWorkbook">True when Excel has an active workbook.</param>
/// <param name="LogAvailable">True when the rolling log exposes an active file path.</param>
/// <param name="SheetInitialised">True when the active workbook has a Gantt sheet that has been initialised (tblGanttData exists).</param>
/// <param name="PlotStartAuto">True when the plot start derives from the data range (R5.1).</param>
/// <param name="PlotFinishAuto">True when the plot finish derives from the data range (R5.1).</param>
/// <param name="PlotStartDate">The explicit plot start date, or empty when automatic (R5.1).</param>
/// <param name="PlotFinishDate">The explicit plot finish date, or empty when automatic (R5.1).</param>
/// <param name="EffectivePlotStartDate">
/// The date the start edit box displays: the explicit date in explicit mode,
/// the derived data range start in automatic mode, or empty when neither is
/// known (fix plan ruling 1 — manual-entry boxes also display the current
/// effective plot date, even when disabled).
/// </param>
/// <param name="EffectivePlotFinishDate">
/// The date the finish edit box displays: the explicit date in explicit mode,
/// the derived data range finish in automatic mode, or empty when neither is
/// known (fix plan ruling 1).
/// </param>
/// <param name="TimeScale">The stored plot time scale (Month/Quarter/Week).</param>
/// <param name="Margin">The stored plot margin preset (Narrow/Normal/Wide/Custom).</param>
/// <param name="MarginCm">The stored custom margin, in centimetres, used only when <paramref name="Margin"/> is <see cref="GanttPlotMargin.Custom"/></param>
/// <param name="Preset">The stored output size preset, or null when the workbook names none.</param>
/// <param name="PlotWidthPt">The plot width the resolver would derive, in points, or null when no preset is known.</param>
/// <param name="PlotHeightPt">The plot height the resolver would derive, in points, or null when no preset is known.</param>
/// <param name="PeriodLabelFormat">
/// The stored period label format: <see cref="GanttPeriodLabelFormat.MM"/> or
/// <see cref="GanttPeriodLabelFormat.MMM"/> on the month scale, the scale's
/// canonical form on the quarter and week scales (R5.12).
/// </param>
/// <param name="DescriptionLabelSelectionAvailable">
/// True when the selection resolves to exactly one table body row (R5.5
/// ruling 6). False disables the description-label dropdown; a collapsed
/// parent row still qualifies (ruling 7) and edits its own cell only.
/// </param>
/// <param name="SelectedDescriptionLabelPosition">
/// The selected row's stored <c>LabelPosition</c> (its
/// description-label position), or null when the cell
/// is blank (Default) or the selection is not a single row.
/// </param>
/// <param name="DescriptionLabelOptions">
/// The selected row's allowed description-label positions in ribbon
/// display order
/// (projected from the catalogue/style — R5.5 D1), or null when the
/// selection is not a single row.
/// </param>
/// <param name="DateLabelSelectionAvailable">
/// True when the selection resolves to exactly one table body row
/// (R5.14). False disables the date-label dropdown; a collapsed parent
/// row still qualifies and edits its own cell only.
/// </param>
/// <param name="SelectedDateLabelPosition">
/// The selected row's stored <c>DateLabelPosition</c> (its
/// date-label choice), or null when the cell is blank (Auto) or the
/// selection is not a single row.
/// </param>
/// <param name="DateLabelOptions">
/// The selected row's allowed date-label choices in ribbon display
/// order (projected from the row's Type — R5.14 D1/D4), or null when
/// the selection is not a single row.
/// </param>
/// <param name="DateLabelPlacementSelectionAvailable">
/// True when the selection resolves to exactly one body row whose
/// Type's date label is positionable — the delineator (owner ruling
/// 2026-10-14). False disables the date-label dropdown for every
/// other type, whose date label selects which date shows.
/// </param>
/// <param name="SelectedDateLabelPlacement">
/// The selected row's stored <c>DateLabelPosition</c> read as a
/// date-label <em>position</em> (the delineator's corner), or null
/// when the cell is blank (Auto) or the selection is not a single
/// positionable row.
/// </param>
/// <param name="DateLabelPlacementOptions">
/// The selected row's allowed date-label positions in ribbon display
/// order (the description label's corner set), or null when the
/// selection is not a single positionable row.
/// </param>
/// <param name="CustomStyleSelectionAvailable">
/// True when the selection resolves to exactly one table body row whose
/// Type carries a style key (R5.6a — Custom Activity, Custom Milestone).
/// False disables the custom-style dropdown; a collapsed parent row
/// still qualifies and edits its own cell only.
/// </param>
/// <param name="SelectedCustomStyleKey">
/// The selected row's stored <c>StyleKey</c>, or null when the cell is
/// blank, names a built-in preset, or the selection is not a single
/// style-key row. Only user-created (non-built-in) keys are offered, so
/// a built-in assignment reads as no match rather than a phantom item.
/// </param>
/// <param name="CustomStyleOptions">
/// The user-created (non-built-in) style keys in catalogue order, or
/// null when the selection is not a single style-key row. Built-in
/// preset keys are never listed: the dropdown assigns custom styles,
/// and the catalogue remains the authority for what counts as built-in.
/// </param>
/// <param name="ConfigurationPresent">
/// True when the active workbook carries a <c>_GanttCreatorConfig</c>
/// worksheet. It separates a first run from a damaged workbook that needs
/// repair (R5.16), and a readable table without configuration is still
/// <c>NeedsRepair</c> (R5.17b).
/// </param>
/// <param name="GanttSheetPresent">
/// True when a worksheet carries the standard Gantt label. A workbook with no
/// configuration and no table but a labelled sheet is recoverable, not a first
/// run, so Initialise cannot create a second Gantt worksheet (R5.17b).
/// </param>
/// <param name="Mode">
/// The Ribbon mode derived from workbook presence, table readability, and
/// configuration presence (R5.16). Drives group visibility only.
/// </param>
/// <param name="CriticalIntervalParentAvailable">
/// True when the current table contains at least one row that can own a
/// Critical Interval (R5.16). The menu item is disabled only when this is
/// false, because the owner's rule is "refuse only when no parent exists".
/// </param>
/// <param name="RowActions">
/// Which structural actions the row under the active cell supports (R5.9b).
/// Every flag mirrors a planner refusal, so a disabled button is always a
/// refusal the command would produce.
/// </param>
/// <param name="ActiveCellInBodyRow">
/// True when the active cell resolves to a table body row (R5.9b). Remove is
/// disabled otherwise, instead of raising a refusal dialog for a header cell.
/// </param>
/// <param name="HasGroups">
/// True when any table row names a parent (R5.16b). Expand all and Collapse all
/// are disabled on a flat table, where they are a quiet no-op. The collapsed
/// state is deliberately not read (plan §5 deferral).
/// </param>
/// <param name="ActiveSheetIsGanttSheet">
/// True when the workbook's active sheet is the worksheet that carries
/// <c>tblGanttData</c> (R5.16b, ADR-0058). Every control outside the recovery
/// surface is disabled otherwise, so a chart command cannot be run while a
/// different visible sheet is active. Defaults to <see langword="true"/> — the
/// fail-open direction, matching <see cref="IsEnabled(string?)"/> — and the
/// state service always supplies the measured fact.
/// </param>
internal sealed record RibbonState(
    bool HasActiveWorkbook,
    bool LogAvailable,
    bool SheetInitialised,
    bool PlotStartAuto,
    bool PlotFinishAuto,
    string PlotStartDate,
    string PlotFinishDate,
    string EffectivePlotStartDate,
    string EffectivePlotFinishDate,
    GanttTimeScale TimeScale,
    GanttPlotMargin Margin,
    double MarginCm,
    SizePreset? Preset,
    double? PlotWidthPt,
    double? PlotHeightPt,
    GanttPeriodLabelFormat PeriodLabelFormat,
    bool DescriptionLabelSelectionAvailable = false,
    GanttLabelPosition? SelectedDescriptionLabelPosition = null,
    IReadOnlyList<GanttLabelPosition>? DescriptionLabelOptions = null,
    bool DateLabelSelectionAvailable = false,
    GanttDateLabelPosition? SelectedDateLabelPosition = null,
    IReadOnlyList<GanttDateLabelPosition>? DateLabelOptions = null,
    bool DateLabelPlacementSelectionAvailable = false,
    GanttLabelPosition? SelectedDateLabelPlacement = null,
    IReadOnlyList<GanttLabelPosition>? DateLabelPlacementOptions = null,
    bool CustomStyleSelectionAvailable = false,
    string? SelectedCustomStyleKey = null,
    IReadOnlyList<string>? CustomStyleOptions = null,
    bool ConfigurationPresent = false,
    bool GanttSheetPresent = false,
    RibbonMode Mode = RibbonMode.FirstRun,
    bool CriticalIntervalParentAvailable = false,
    RowActionAvailability RowActions = default,
    bool ActiveCellInBodyRow = false,
    bool HasGroups = false,
    bool ActiveSheetIsGanttSheet = true)
{
    /// <summary>
    /// The snapshot before any successful capture. No fact is known, so every
    /// gated control starts disabled and is enabled by the first refresh
    /// (docs/03-ROADMAP.md R1.5). The scale and margin default to the catalogue
    /// defaults so the dropdowns and combobox show a selection even before the
    /// workbook is read (R5.2); the width and height displays show nothing,
    /// because there is no preset to measure against yet.
    /// </summary>
    internal static RibbonState Initial { get; } = new(
        false,
        false,
        false,
        true,
        true,
        string.Empty,
        string.Empty,
        string.Empty,
        string.Empty,
        GanttTimeScale.Month,
        GanttPlotMargin.Normal,
        GanttPlotMargins.DefaultCustomCm,
        null,
        null,
        null,
        GanttPeriodLabelFormat.MMM);

    /// <summary>
    /// Gets whether the control with <paramref name="controlId"/> is enabled
    /// under this snapshot: the pure truth table behind the Ribbon's
    /// <c>getEnabled</c> callback.
    /// </summary>
    /// <param name="controlId">The Ribbon control ID, or null when it could not be probed.</param>
    /// <returns>
    /// The control's enabled state. Unknown, null, or whitespace IDs are
    /// <see langword="true"/> (fail-open): a permanently grey button is a worse
    /// failure than a briefly enabled one, and the command boundary still
    /// validates at execution time.
    /// </returns>
    internal bool IsEnabled(string? controlId) => controlId switch
    {
        // R5.16b (ADR-0058): Initialise, Repair, Validate, Diagnostics and Open log
        // are the recovery surface and stay available on any sheet. Everything
        // else also requires the Gantt sheet to be the active sheet, so a chart
        // command is never offered against a different visible sheet.
        RibbonControlIds.Diagnostics => HasActiveWorkbook,
        RibbonControlIds.OpenLog => LogAvailable,
        RibbonControlIds.InitialiseSheet => HasActiveWorkbook,
        RibbonControlIds.ValidateSheet => HasActiveWorkbook,
        RibbonControlIds.RepairConfig => HasActiveWorkbook,
        RibbonControlIds.RefreshSheet =>
            HasActiveWorkbook && SheetInitialised && ActiveSheetIsGanttSheet,
        RibbonControlIds.Warnings =>
            HasActiveWorkbook && SheetInitialised && ActiveSheetIsGanttSheet,
        RibbonControlIds.AddActivity =>
            HasActiveWorkbook && SheetInitialised && ActiveSheetIsGanttSheet,
        RibbonControlIds.AddMilestone =>
            HasActiveWorkbook && SheetInitialised && ActiveSheetIsGanttSheet,
        RibbonControlIds.AddDelineator =>
            HasActiveWorkbook && SheetInitialised && ActiveSheetIsGanttSheet,
        // R5.16: a Critical Interval must be attached to a parent, so the item is
        // disabled when no row can own one. It stays VISIBLE: the owner asked for
        // a constraint that is explained in place, not a hidden action.
        RibbonControlIds.CriticalIntervalMenuItem =>
            HasActiveWorkbook && SheetInitialised && ActiveSheetIsGanttSheet
            && CriticalIntervalParentAvailable,
        // R5.9b: the structural buttons follow the row under the active cell.
        // The flags come from RowActionAvailability, whose rules mirror the
        // planner refusals, so a grey button is always a refusal the command
        // would produce and never a condition it can satisfy.
        RibbonControlIds.PromoteToParent =>
            HasActiveWorkbook && SheetInitialised && ActiveSheetIsGanttSheet
            && RowActions.CanPromote,
        RibbonControlIds.DemoteToChild =>
            HasActiveWorkbook && SheetInitialised && ActiveSheetIsGanttSheet
            && RowActions.CanDemote,
        RibbonControlIds.MoveUp =>
            HasActiveWorkbook && SheetInitialised && ActiveSheetIsGanttSheet
            && RowActions.CanMoveUp,
        RibbonControlIds.MoveDown =>
            HasActiveWorkbook && SheetInitialised && ActiveSheetIsGanttSheet
            && RowActions.CanMoveDown,
        // R5.5 ruling 6: the description-label dropdown is the first
        // selection-gated control — it enables only for exactly one
        // resolved body row, so a header/multi/no selection disables it
        // and the command still re-checks at execution time (ruling 1).
        RibbonControlIds.DescriptionLabel =>
            HasActiveWorkbook && SheetInitialised && ActiveSheetIsGanttSheet
            && DescriptionLabelSelectionAvailable,
        // R5.14: the date-label dropdown is gated on the same single-body-row
        // selection facts as the description-label dropdown (the capability is
        // type-only, so the options projection needs no style registry).
        // Owner ruling 2026-10-14: the dropdown carries one of two
        // vocabularies — which-date choices for dated types, positions
        // for the delineator — so it enables when either context
        // resolved.
        RibbonControlIds.DateLabel =>
            HasActiveWorkbook && SheetInitialised && ActiveSheetIsGanttSheet
            && (DateLabelSelectionAvailable || DateLabelPlacementSelectionAvailable),
        // R5.6a: the custom-style dropdown is gated on the same single-body-row
        // selection facts as the description-label dropdown, narrowed to
        // style-key types (Custom Activity / Custom Milestone) AND at least
        // one user-created style key existing: an enabled dropdown with zero
        // items reads as broken, so a workbook with no custom styles leaves
        // it greyed until Theme settings creates one.
        RibbonControlIds.CustomStyle =>
            HasActiveWorkbook && SheetInitialised && ActiveSheetIsGanttSheet
            && CustomStyleSelectionAvailable
            && CustomStyleOptions is { Count: > 0 },
        // R5.9b: Remove needs a body row under the active cell; a header or an
        // off-table cell leaves it grey instead of raising a refusal dialog.
        RibbonControlIds.Remove =>
            HasActiveWorkbook && SheetInitialised && ActiveSheetIsGanttSheet
            && ActiveCellInBodyRow,
        // R5.4: the collapse-all and expand-all buttons are the Ribbon surface
        // over the native Excel outline. R5.16b adds the HasGroups fact: with no
        // row naming a parent they are a quiet no-op, so they are greyed instead.
        // The collapsed state is still not read — it would be a COM outline
        // query per refresh for a condition the command already handles.
        RibbonControlIds.ExpandAllGroups =>
            HasActiveWorkbook && SheetInitialised && ActiveSheetIsGanttSheet && HasGroups,
        RibbonControlIds.CollapseAllGroups =>
            HasActiveWorkbook && SheetInitialised && ActiveSheetIsGanttSheet && HasGroups,
        RibbonControlIds.PlotStartAuto =>
            HasActiveWorkbook && SheetInitialised && ActiveSheetIsGanttSheet,
        RibbonControlIds.PlotFinishAuto =>
            HasActiveWorkbook && SheetInitialised && ActiveSheetIsGanttSheet,
        RibbonControlIds.PlotStartDate =>
            HasActiveWorkbook && SheetInitialised && ActiveSheetIsGanttSheet && !PlotStartAuto,
        RibbonControlIds.PlotFinishDate =>
            HasActiveWorkbook && SheetInitialised && ActiveSheetIsGanttSheet && !PlotFinishAuto,
        RibbonControlIds.PlotTimeScale =>
            HasActiveWorkbook && SheetInitialised && ActiveSheetIsGanttSheet,
        RibbonControlIds.Margin =>
            HasActiveWorkbook && SheetInitialised && ActiveSheetIsGanttSheet,
        // R5.16a: the four preset buttons are one dropdown. The width/height
        // readouts are outputs, never inputs, so they are labelControls and
        // never appear in this table.
        RibbonControlIds.SizeDropdown =>
            HasActiveWorkbook && SheetInitialised && ActiveSheetIsGanttSheet,
        RibbonControlIds.ThemeSettings =>
            HasActiveWorkbook && SheetInitialised && ActiveSheetIsGanttSheet,
        RibbonControlIds.ChartTitle =>
            HasActiveWorkbook && SheetInitialised && ActiveSheetIsGanttSheet,
        _ => true,
    };

    /// <summary>
    /// Gets whether the group with <paramref name="groupId"/> is visible under
    /// this snapshot: the pure truth table behind the Ribbon's
    /// <c>getVisible</c> callback (R5.16).
    /// </summary>
    /// <param name="groupId">The Ribbon group ID, or null when it could not be probed.</param>
    /// <returns>
    /// The group's visibility. Unlike <see cref="IsEnabled"/>, visibility is
    /// FAIL-CLOSED: an unknown or null id is hidden, because showing an
    /// unintended group is the failure R5.16 removes. Diagnostics is the one
    /// group visible in every mode (owner ruling 2026-11-10).
    /// </returns>
    internal bool IsVisible(string? groupId) => groupId switch
    {
        RibbonControlIds.DiagnosticsGroup => true,
        RibbonControlIds.GetStartedGroup => Mode == RibbonMode.FirstRun,
        RibbonControlIds.DataGroup
            or RibbonControlIds.OutlineGroup
            or RibbonControlIds.LabelsGroup
            or RibbonControlIds.StyleThemeGroup
            or RibbonControlIds.PlotAreaGroup
            or RibbonControlIds.PlotSizeGroup
            or RibbonControlIds.SheetGroup => Mode != RibbonMode.FirstRun,
        _ => false,
    };
}
