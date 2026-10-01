using GanttCreator.Core;
using GanttCreator.Core.Scene;

namespace GanttCreator.Office;

/// <summary>Why a <see cref="ISceneBuildRequestFactory"/> could not produce a request.</summary>
/// <remarks>
/// Every reason is a typed, expected outcome rather than an exception
/// (<c>docs/02-ARCHITECTURE.md</c> "Error handling"). A caller that receives one of
/// these has an actionable message and no chart was built.
/// </remarks>
public enum SceneBuildRequestRefusal
{
    /// <summary>The application object or active workbook was absent.</summary>
    NoActiveWorkbook = 0,

    /// <summary>The Gantt worksheet or its table was missing.</summary>
    TableMissing = 1,

    /// <summary>The configuration sheet could not be read.</summary>
    ConfigurationUnreadable = 2,

    /// <summary>The live panel-grid measurement refused.</summary>
    MeasurementRefused = 3,

    /// <summary>A configured value could not be parsed into its typed form.</summary>
    InvalidSetting = 4,

    /// <summary>
    /// The configured size preset is not a known preset key.
    /// </summary>
    /// <remarks>
    /// A distinct reason from <see cref="InvalidSetting"/> on purpose. A bad preset
    /// key is a far more common and far more recoverable mistake than a malformed
    /// number, and the two deserve different messages: one names the key that is not
    /// available, the other names a value that will not parse.
    /// </remarks>
    UnknownSizePreset = 5,

    /// <summary>
    /// The plot could not be resolved from the preset and the measured panel.
    /// </summary>
    /// <remarks>
    /// The live instance is a text panel too wide for the chosen paper size, which
    /// <see cref="PlotGeometryResolver"/> refuses rather than shrinking the panel.
    /// </remarks>
    PlotGeometryRefused = 6,

    /// <summary>The configured events carry no renderable date range.</summary>
    NoPlotRange = 7,
}

/// <summary>
/// The typed outcome of one scene-request build: the request, or why there is none.
/// </summary>
/// <param name="Request">The request on success, or <see langword="null"/>.</param>
/// <param name="Refusal">The refusal reason, or <see langword="null"/> on success.</param>
/// <param name="Message">A user-facing explanation, present on a refusal.</param>
public sealed record SceneBuildRequestOutcome(
    SceneBuildRequest? Request,
    SceneBuildRequestRefusal? Refusal,
    string? Message)
{
    /// <summary>Gets whether a request was produced.</summary>
    public bool Succeeded => Request is not null;

    /// <summary>Creates a successful outcome.</summary>
    /// <param name="request">The produced request.</param>
    /// <returns>The successful outcome.</returns>
    public static SceneBuildRequestOutcome Ok(SceneBuildRequest request) => new(request, null, null);

    /// <summary>Creates a refusal with a user-facing message.</summary>
    /// <param name="refusal">Why no request was produced.</param>
    /// <param name="message">The explanation shown to the user.</param>
    /// <returns>The refusal outcome.</returns>
    public static SceneBuildRequestOutcome Refused(SceneBuildRequestRefusal refusal, string message) =>
        new(null, refusal, message);
}

/// <summary>
/// Owns the translation of live workbook and configuration state into the
/// <see cref="SceneBuildRequest"/> that <c>SceneBuilder</c> consumes (R4.8A D5).
/// </summary>
/// <remarks>
/// <para>
/// <b>Why this is a port and not inline logic.</b> D5 requires that visible-row
/// geometry, projections, plot bounds, the size preset, the date range, the scale,
/// the style registry, the metrics, and the composition profile are all decided
/// <em>here</em> and nowhere else. The alternative — a Ribbon callback or a renderer
/// deriving them — is how a second authority appears: two places computing plot
/// bounds, or one place resolving the preset and another re-deriving the range from
/// it. An architecture test enforces the single-caller rule; this port is what it
/// enforces it over.
/// </para>
/// <para>
/// <b>Why the factory is separately testable.</b> It takes no dependency on a
/// worksheet, a shape port, or a command. A test supplies rows, a settings map, a
/// style registry, and a measured grid, and asserts the exact request produced. The
/// parts it cannot compute for itself — the measured grid and the plot bounds — are
/// dependencies it is given, not things it guesses.
/// </para>
/// <para>
/// The surface is primitive: no interop type crosses it, so the AddIn compilation
/// never names <c>Microsoft.Office.Interop.Excel</c>.
/// </para>
/// </remarks>
public interface ISceneBuildRequestFactory
{
    /// <summary>Builds the scene request for the current workbook state.</summary>
    /// <param name="events">The validated events to render.</param>
    /// <param name="settings">The effective settings key/value map.</param>
    /// <param name="registry">The named-style registry to resolve each event against.</param>
    /// <param name="grid">The measured live panel cell grid.</param>
    /// <returns>The request, or a typed refusal carrying a user-facing message.</returns>
    SceneBuildRequestOutcome Create(
        IReadOnlyList<GanttEvent> events,
        IReadOnlyDictionary<string, string> settings,
        GanttStyleRegistry registry,
        PanelCellGrid grid);
}
