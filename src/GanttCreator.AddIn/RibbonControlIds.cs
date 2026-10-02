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
}
