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

    /// <summary>The Repair configuration button (<c>getEnabled</c> driven by workbook presence).</summary>
    internal const string RepairConfig = "btnRepairConfig";

    /// <summary>The Add activity button (<c>getEnabled</c> driven by workbook presence).</summary>
    internal const string AddActivity = "btnAddActivity";

    /// <summary>The Add milestone button (<c>getEnabled</c> driven by workbook presence).</summary>
    internal const string AddMilestone = "btnAddMilestone";

    /// <summary>The Add delineator button (<c>getEnabled</c> driven by workbook presence).</summary>
    internal const string AddDelineator = "btnAddDelineator";

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

    /// <summary>The plot width display (always disabled; read-only output).</summary>
    internal const string PlotWidth = "edtPlotWidth";

    /// <summary>The plot height display (always disabled; read-only output).</summary>
    internal const string PlotHeight = "edtPlotHeight";

    /// <summary>The A4 portrait preset button (<c>getEnabled</c> driven by workbook + initialised).</summary>
    internal const string PresetA4Portrait = "btnPresetA4Portrait";

    /// <summary>The A4 landscape preset button (<c>getEnabled</c> driven by workbook + initialised).</summary>
    internal const string PresetA4Landscape = "btnPresetA4Landscape";

    /// <summary>The Presentation 16:9 preset button (<c>getEnabled</c> driven by workbook + initialised).</summary>
    internal const string PresetPresentation16x9 = "btnPresetPresentation16x9";

    /// <summary>The Presentation 4:3 preset button (<c>getEnabled</c> driven by workbook + initialised).</summary>
    internal const string PresetPresentation4x3 = "btnPresetPresentation4x3";
}
