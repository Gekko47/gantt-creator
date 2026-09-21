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
}
